import pandas as pd
from difflib import SequenceMatcher
from typing import List, Dict, Optional


def match_descriptions(
        lines: List[Dict],
        df_codes: pd.DataFrame,
        min_score: float = 0.18,
        gemini_threshold: float = 0.30,
        use_gemini: bool = True,
        gemini_api_key: Optional[str] = None
) -> List[Dict]:
    """
    Potriveste fiecare linie din factura cu cel mai apropiat cod din Excel.
    Foloseste Gemini AI pentru cazuri cu confidence scazut (scor < gemini_threshold).

    Args:
        lines: Lista de linii din factura de potrivit
        df_codes: DataFrame cu coduri de referinta
        min_score: Scor minim fuzzy pentru a considera (default: 0.18)
        gemini_threshold: Daca cel mai bun rezultat < aceasta, foloseste Gemini AI (default: 0.30)
        use_gemini: Activeaza analiza Gemini AI (default: True)
        gemini_api_key: Cheia API Gemini (optional)

    Returns:
        Lista de linii cu coduri potrivite si scoruri de confidence
    """
    cols = list(df_codes.columns)
    if len(cols) < 2:
        raise ValueError("Fisierul Excel trebuie sa aiba cel putin doua coloane (cod si denumire).")

    code_col, desc_col = cols[0], cols[1]

    gemini_matcher = None
    if use_gemini:
        try:
            from gemini_matcher import GeminiMatcher
            gemini_matcher = GeminiMatcher(api_key=gemini_api_key)
            print(f"[OK] Gemini AI activat (prag: {gemini_threshold})")
        except ImportError:
            print("[ATENTIE] google-generativeai nu este instalat. Ruleaza: pip install google-generativeai")
            use_gemini = False
        except ValueError as e:
            print(f"[ATENTIE] Gemini dezactivat - {e}")
            use_gemini = False

    results = []
    gemini_calls = 0

    for line in lines:
        input_desc = str(line.get("description", "")).lower()

        matches = []
        for _, row in df_codes.iterrows():
            score = SequenceMatcher(None, input_desc, str(row[desc_col]).lower()).ratio()
            if score >= min_score:
                matches.append({
                    "matched_code": row[code_col],
                    "matched_description": row[desc_col],
                    "score": round(score, 4)
                })

        matches.sort(key=lambda x: x["score"], reverse=True)

        if not matches:
            line["matched_code"] = None
            line["matched_description"] = None
            line["score"] = 0.0
            line["status"] = "fara_potriviri"
            results.append(line)
            continue

        best_match = matches[0]

        if use_gemini and gemini_matcher and best_match["score"] < gemini_threshold:
            try:
                top_candidates = matches[:5]
                gemini_result = gemini_matcher.analyze_candidates(
                    product_description=line.get("description", ""),
                    candidates=top_candidates
                )

                line["matched_code"] = gemini_result["matched_code"]
                line["matched_description"] = gemini_result["matched_description"]
                line["score"] = gemini_result["confidence"]
                line["status"] = gemini_result["status"]
                line["reasoning"] = gemini_result.get("reasoning", "")
                line["fuzzy_score"] = best_match["score"]  # Keep original fuzzy score

                gemini_calls += 1
                status_icon = "[X]" if gemini_result["matched_code"] is None else "[OK]"
                print(f"  [AI] Analiza Gemini #{gemini_calls} {status_icon}: {line.get('description', '')[:50]}... -> {line['matched_code']}")

            except Exception as e:
                print(f"  [EROARE] Gemini: {e}. Folosesc matching-ul fuzzy.")
                line.update(best_match)
                line["status"] = "eroare_gemini_fallback"
        else:
            line.update(best_match)
            line["status"] = "potrivit_fuzzy" if best_match["score"] >= gemini_threshold else "fuzzy_confidence_scazut"

        results.append(line)

    print(f"\n[SUMAR] Potriviri:")
    print(f"  Total linii: {len(results)}")
    print(f"  Fuzzy confidence ridicat: {sum(1 for r in results if r.get('status') == 'potrivit_fuzzy')}")
    print(f"  Analizate de Gemini: {gemini_calls}")
    print(f"    Acceptate: {sum(1 for r in results if r.get('status') == 'gemini_analyzed' and r.get('matched_code'))}")
    print(f"    Respinse (null): {sum(1 for r in results if r.get('status') == 'gemini_no_match')}")
    print(f"  Fara potriviri: {sum(1 for r in results if r.get('status') == 'fara_potriviri')}")

    return results
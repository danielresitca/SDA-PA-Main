# -*- coding: utf-8 -*-
"""
Gemini API matcher pentru cazuri cu confidence scazut in matching-ul fuzzy.
Analizeaza top 5 candidati si selecteaza cel mai potrivit cod sau niciun cod.
"""

import os
import json
from typing import List, Dict, Optional
import google.generativeai as genai


class GeminiMatcher:
    """Matcher inteligent folosind Gemini API pentru cazuri cu confidence scazut."""

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialializeaza matcher-ul Gemini.

        Args:
            api_key: Cheia API Gemini. Daca None, citeste din variabila GEMINI_API_KEY.
        """
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")

        if not self.api_key:
            raise ValueError(
                "Cheia API Gemini nu a fost gasita. Furnizeaza-o ca argument sau seteaza variabila GEMINI_API_KEY."
            )

        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel('gemini-pro')

    def analyze_candidates(
            self,
            product_description: str,
            candidates: List[Dict],
            min_candidates: int = 5
    ) -> Dict:
        """
        Analizeaza candidatii din matching-ul fuzzy folosind Gemini si selecteaza cel mai bun.

        Args:
            product_description: Denumirea produsului din factura
            candidates: Lista de candidati din matching-ul fuzzy
            min_candidates: Numarul minim de candidati de trimis (default: 5)

        Returns:
            Dict cu matched_code, matched_description, confidence si reasoning
        """
        candidates_to_analyze = candidates[:max(min_candidates, len(candidates))]

        if not candidates_to_analyze:
            return self._create_no_match_result(product_description)

        prompt = self._build_analysis_prompt(product_description, candidates_to_analyze)

        try:
            response = self.model.generate_content(prompt)

            result = self._parse_gemini_response(response.text, candidates_to_analyze)
            result["ai_method"] = "gemini"

            return result

        except Exception as e:
            print(f"Eroare API Gemini: {e}")
            return self._create_fallback_result(candidates_to_analyze[0], str(e))

    def _build_analysis_prompt(self, description: str, candidates: List[Dict]) -> str:
        """Build the analysis prompt for Gemini in Romanian."""

        candidates_text = "\n".join([
            f"{i + 1}. Cod: {c['matched_code']}\n   Denumire: {c['matched_description']}\n   Scor fuzzy: {c['score']:.2f}"
            for i, c in enumerate(candidates)
        ])

        prompt = f"""Esti expert in potrivirea descrierilor de produse din facturi cu o baza de date de referinta.

PRODUS DIN FACTURA:
"{description}"

TOP 5 CODURI CANDIDATE (alese de algoritmul fuzzy):
{candidates_text}

SARCINA TA:
Analizeaza denumirea produsului din factura si ALEGE cel mai potrivit cod din cele 5 optiuni de mai sus.

REGULI STRICTE:
1. Poti alege DOAR unul din cele 5 coduri de mai sus
2. Daca NICIUN cod nu se potriveste produsului, returneaza "selected_code": null
3. Nu inventa coduri noi - foloseste EXACT codul din lista sau null

Criterii de evaluare:
- Categoria de produs si materiale
- Specificatii tehnice (daca exista)
- Utilizarea finala a produsului
- Termeni specifici industriei

RASPUNDE DOAR cu un obiect JSON in acest format EXACT (fara markdown, fara text extra):
{{
  "selected_code": "codul ales SAU null",
  "confidence": 0.85,
  "reasoning": "Explicatie scurta in romana de ce ai ales acest cod (max 100 cuvinte)"
}}

Niveluri de confidence:
- 0.9-1.0: Potrivire foarte sigura
- 0.7-0.9: Potrivire buna cu mica incertitudine
- 0.5-0.7: Potrivire acceptabila, recomand verificare manuala
- Sub 0.5: Incertitudine mare, necesita verificare manuala
- 0.0: Niciun cod nu se potriveste (selected_code: null)"""

        return prompt

    def _parse_gemini_response(self, response_text: str, candidates: List[Dict]) -> Dict:
        """Parse Gemini's JSON response."""

        try:
            response_text = response_text.strip()

            if response_text.startswith("```"):
                lines = response_text.split("\n")
                json_lines = []
                in_json = False
                for line in lines:
                    if line.strip().startswith("{"):
                        in_json = True
                    if in_json:
                        json_lines.append(line)
                    if line.strip().endswith("}"):
                        break
                response_text = "\n".join(json_lines)

            gemini_result = json.loads(response_text)

            selected_code = gemini_result.get("selected_code")
            confidence = float(gemini_result.get("confidence", 0.5))
            reasoning = gemini_result.get("reasoning", "Nicio explicatie furnizata")

            if selected_code is None or selected_code == "null":
                return {
                    "matched_code": None,
                    "matched_description": None,
                    "confidence": confidence,
                    "reasoning": reasoning,
                    "status": "gemini_no_match"
                }

            selected_candidate = next(
                (c for c in candidates if c["matched_code"] == selected_code),
                None
            )

            if selected_candidate is None:
                print(f"[ATENTIE] Gemini a ales un cod invalid: {selected_code}")
                return self._create_fallback_result(
                    candidates[0],
                    f"AI a ales un cod invalid ({selected_code}) care nu era in lista"
                )

            return {
                "matched_code": selected_candidate["matched_code"],
                "matched_description": selected_candidate["matched_description"],
                "confidence": confidence,
                "reasoning": reasoning,
                "status": "gemini_analyzed"
            }

        except (json.JSONDecodeError, KeyError, ValueError) as e:
            print(f"Eroare la parsarea raspunsului Gemini: {e}")
            print(f"Raspuns brut: {response_text[:200]}")
            return self._create_fallback_result(candidates[0], "Nu s-a putut parsa raspunsul AI")

    def _create_fallback_result(self, best_candidate: Dict, error_msg: str) -> Dict:
        """Create fallback result using best fuzzy match."""
        return {
            "matched_code": best_candidate["matched_code"],
            "matched_description": best_candidate["matched_description"],
            "confidence": best_candidate["score"],
            "reasoning": f"Analiza Gemini a esuat: {error_msg}. Folosesc cel mai bun rezultat fuzzy.",
            "status": "gemini_fallback"
        }

    def _create_no_match_result(self, description: str) -> Dict:
        """Create result when no candidates are available."""
        return {
            "matched_code": None,
            "matched_description": None,
            "confidence": 0.0,
            "reasoning": "Niciun candidat disponibil din matching-ul fuzzy",
            "status": "no_match"
        }


# Functie convenabila pentru folosire rapida
def analyze_with_gemini(
        product_description: str,
        candidates: List[Dict],
        api_key: Optional[str] = None
) -> Dict:
    """
    Functie rapida pentru a analiza candidatii cu Gemini.

    Args:
        product_description: Denumirea produsului din factura
        candidates: Lista de candidati din matching-ul fuzzy
        api_key: Cheia API Gemini (optional)

    Returns:
        Rezultatul analizei cu codul selectat si explicatie
    """
    matcher = GeminiMatcher(api_key=api_key)
    return matcher.analyze_candidates(product_description, candidates)
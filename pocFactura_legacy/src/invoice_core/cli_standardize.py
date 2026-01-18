# -*- coding: utf-8 -*-
# CLI entry point pentru standardizare cu suport Gemini AI
import json
import argparse
import pandas as pd
import os
import sys
from matcher import match_descriptions
from ubl_lines import extract_lines

# Incearca sa importe configurarea
try:
    from config import GEMINI_API_KEY as CONFIG_API_KEY
except ImportError:
    CONFIG_API_KEY = None


def main():
    # Fix encoding on Windows
    if sys.platform == 'win32':
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')

    parser = argparse.ArgumentParser(
        description="Standardizare linii factura folosind matching fuzzy si Gemini AI"
    )
    parser.add_argument("--xml", required=True, help="Calea catre fisierul XML cu factura")
    parser.add_argument("--codes-xlsx", required=True, help="Calea catre fisierul Excel cu coduri de referinta")
    parser.add_argument("--out-lines", default="lines_raw.json", help="Fisier de iesire pentru linii brute")
    parser.add_argument("--out-standard", default="lines_standardized.json",
                        help="Fisier de iesire pentru linii standardizate")
    parser.add_argument("--out-csv", default="lines_standardized.csv", help="Fisier CSV de iesire")
    parser.add_argument("--min-score", type=float, default=0.18, help="Scor minim fuzzy match")
    parser.add_argument("--gemini-threshold", type=float, default=0.30,
                        help="Foloseste Gemini AI cand scorul fuzzy e sub acest prag (default: 0.30)")
    parser.add_argument("--no-gemini", action="store_true",
                        help="Dezactiveaza analiza Gemini AI (foloseste doar matching fuzzy)")
    parser.add_argument("--gemini-api-key", type=str, default=None,
                        help="Cheia API Gemini (sau seteaza in config.py sau variabila GEMINI_API_KEY)")

    args = parser.parse_args()

    print("=" * 60)
    print("Instrument de Standardizare Facturi")
    print("=" * 60)

    # Extract lines from XML
    print(f"\n[1/4] Extragere linii din: {args.xml}")
    try:
        lines = extract_lines(args.xml)
        print(f"  [OK] Gasite {len(lines)} linii")
    except Exception as e:
        print(f"  [EROARE] {e}")
        return 1

    # Load codes database
    print(f"\n[2/4] Incarcare coduri din: {args.codes_xlsx}")
    try:
        df_codes = pd.read_excel(args.codes_xlsx)
        print(f"  [OK] Incarcate {len(df_codes)} coduri")
    except Exception as e:
        print(f"  [EROARE] {e}")
        return 1

    # Get Gemini API key - prioritate: argument -> config.py -> variabila mediu
    gemini_api_key = args.gemini_api_key or CONFIG_API_KEY or os.getenv("GEMINI_API_KEY")

    # Match descriptions
    print(f"\n[3/4] Potrivire denumiri...")
    print(f"  Prag fuzzy: {args.min_score}")
    print(f"  Prag Gemini: {args.gemini_threshold}")

    if args.no_gemini:
        print(f"  Gemini AI: Dezactivat")
    else:
        if gemini_api_key:
            print(f"  Gemini AI: Activat (cheie gasita)")
        else:
            print(f"  Gemini AI: Dezactivat (nicio cheie API gasita)")
            print(f"    Adauga cheia in config.py sau foloseste --gemini-api-key")

    try:
        standardized = match_descriptions(
            lines,
            df_codes,
            min_score=args.min_score,
            gemini_threshold=args.gemini_threshold,
            use_gemini=not args.no_gemini,
            gemini_api_key=gemini_api_key
        )
    except Exception as e:
        print(f"  [EROARE] {e}")
        return 1

    # Save results
    print(f"\n[4/4] Salvare rezultate...")

    try:
        # Save JSON
        with open(args.out_standard, "w", encoding="utf-8") as f:
            json.dump({"lines": standardized}, f, indent=2, ensure_ascii=False)
        print(f"  [OK] JSON: {args.out_standard}")

        # Save CSV
        df_output = pd.DataFrame(standardized)
        df_output.to_csv(args.out_csv, index=False, encoding="utf-8")
        print(f"  [OK] CSV: {args.out_csv}")
    except Exception as e:
        print(f"  [EROARE] {e}")
        return 1

    print("\n" + "=" * 60)
    print("Procesare completa!")
    print("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
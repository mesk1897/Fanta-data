import os
import sys
import json
import requests
import pandas as pd

# URL dell'export ufficiale Fantacalcio
url = "https://www.fantacalcio.it/servizi/Excel/Quotazioni_Fantacalcio_Stagione_2026_27.xlsx"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

print(f"Connessione a: {url}")

try:
    response = requests.get(url, headers=headers, timeout=30)
    print(f"Status HTTP: {response.status_code}")
    print(f"Content-Type: {response.headers.get('Content-Type')}")
    print(f"Dimensione scaricata: {len(response.content)} bytes")

    # Salva quanto scaricato
    with open("quotazioni.xlsx", "wb") as f:
        f.write(response.content)

    # Ispezione dei primi byte
    with open("quotazioni.xlsx", "rb") as f:
        primi_byte = f.read(250)
        print(f"Primi byte ricevuti: {primi_byte}")

    # Verifica se c'è stato un errore di rete (es. 403 o 404)
    if response.status_code != 200:
        print(f"\nERRORE HTTP {response.status_code}: Il server non ha inviato il file.")
        sys.exit(1)

    # Verifica se è una pagina web di errore o captcha invece di un file Excel
    if primi_byte.startswith(b"<!DOCTYPE") or primi_byte.startswith(b"<html") or b"<head>" in primi_byte:
        print("\nBLOCCO RILEVATO: Il server ha risposto con una pagina HTML (possibile protezione anti-bot o link non valido) anziché con il file Excel.")
        sys.exit(1)

    # Lettura del file Excel con motore openpyxl
    df = pd.read_excel("quotazioni.xlsx", engine="openpyxl", skiprows=1)

    players = []
    for _, row in df.iterrows():
        try:
            nome = str(row.get("Nome", "")).strip()
            if not nome or nome.lower() == "nan":
                continue
            players.append({
                "id": int(row.get("Id", 0)) if pd.notnull(row.get("Id")) else 0,
                "ruolo": str(row.get("R", "")).strip(),
                "nome": nome,
                "squadra": str(row.get("Squadra", "")).strip(),
                "quotazione": int(row.get("Qt.A", 1)) if pd.notnull(row.get("Qt.A")) else 1,
                "fantaMedia": float(row.get("Fm", 0)) if pd.notnull(row.get("Fm")) else 0.0,
                "mediaVoto": float(row.get("Mv", 0)) if pd.notnull(row.get("Mv")) else 0.0,
                "gol": int(row.get("Gf", 0)) if pd.notnull(row.get("Gf")) else 0,
                "assist": int(row.get("Ass", 0)) if pd.notnull(row.get("Ass")) else 0,
                "ammonizioni": int(row.get("Amm", 0)) if pd.notnull(row.get("Amm")) else 0,
                "espulsioni": int(row.get("Esp", 0)) if pd.notnull(row.get("Esp")) else 0,
                "status": "disponibile"
            })
        except Exception:
            continue

    with open("dati_serie_a.json", "w", encoding="utf-8") as f:
        json.dump(players, f, ensure_ascii=False, indent=2)

    print(f"\nCompletato: salvati {len(players)} giocatori in dati_serie_a.json.")

except Exception as e:
    print(f"\nErrore durante l'elaborazione: {e}")
    raise e

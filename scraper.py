import os
import sys
import json
import requests
import pandas as pd

# Creazione di una sessione HTTP per conservare i cookie come un vero browser
session = requests.Session()

# Header completi da browser desktop moderno
headers_browser = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
    "Sec-Ch-Ua": '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1"
}

session.headers.update(headers_browser)

url_pagina = "https://www.fantacalcio.it/quotazioni-fantacalcio"
url_excel = "https://www.fantacalcio.it/servizi/Excel/Quotazioni_Fantacalcio_Stagione_2026_27.xlsx"

try:
    print(f"1. Visita alla pagina delle quotazioni per ottenere i cookie di sessione...")
    res_page = session.get(url_pagina, timeout=30)
    print(f"   Risposta pagina: HTTP {res_page.status_code}")

    print(f"\n2. Richiesta download del file Excel con Referer impostato...")
    headers_download = {
        "Referer": url_pagina,
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors"
    }

    response = session.get(url_excel, headers=headers_download, timeout=30)
    print(f"   Status HTTP: {response.status_code}")
    print(f"   Content-Type ricevuto: {response.headers.get('Content-Type')}")
    print(f"   Dimensione scaricata: {len(response.content)} bytes")

    # Salvataggio su disco
    with open("quotazioni.xlsx", "wb") as f:
        f.write(response.content)

    primi_byte = response.content[:10]
    print(f"   Primi byte: {primi_byte}")

    # Verifica: i file Excel .xlsx iniziano SEMPRE con la firma binaria b'PK\x03\x04'
    if not primi_byte.startswith(b"PK\x03\x04"):
        print("\nBLOCCO ANCORA ATTIVO: Il server ha risposto con codice 200 ma il contenuto è HTML (protezione attiva).")
        sys.exit(1)

    print("\nFile Excel autentico confermato! Inizio elaborazione...")
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

    print(f"\nCompletato con successo! Generati {len(players)} giocatori in dati_serie_a.json.")

except Exception as e:
    print(f"\nErrore inaspettato: {e}")
    sys.exit(1)

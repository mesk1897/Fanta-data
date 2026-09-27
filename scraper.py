import os
import sys
import json
import pandas as pd
from curl_cffi import requests

# Sessione che simula Chrome su Windows
session = requests.Session(impersonate="chrome120")

url_pagina = "https://www.fantacalcio.it/quotazioni-fantacalcio"
url_excel = "https://www.fantacalcio.it/api/v1/Excel/prices/21/1"

try:
    print("1. Visita alla pagina delle quotazioni...")
    res_page = session.get(url_pagina, timeout=30)
    print(f"   Risposta pagina: HTTP {res_page.status_code}")

    print(f"\n2. Richiesta download all'endpoint API reale:\n   {url_excel}")
    headers_download = {
        "Referer": url_pagina,
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }

    response = session.get(url_excel, headers=headers_download, timeout=30)
    print(f"   Status HTTP: {response.status_code}")
    print(f"   Content-Type ricevuto: {response.headers.get('Content-Type')}")
    print(f"   Dimensione scaricata: {len(response.content)} bytes")

    # Salvataggio temporaneo per analisi
    with open("quotazioni.xlsx", "wb") as f:
        f.write(response.content)

    primi_byte = response.content[:10]
    print(f"   Primi byte: {primi_byte}")

    # Verifica firma standard file Excel (b'PK\x03\x04')
    if not primi_byte.startswith(b"PK\x03\x04"):
        print("\nRISPOSTA NON VALIDA (Richiede autenticazione o token):")
        try:
            print("Estratto risposta server:\n", response.text[:400])
        except Exception:
            pass
        sys.exit(1)

    print("\nFile Excel autentico confermato! Inizio elaborazione...")
    df = pd.read_excel("quotazioni.xlsx", engine="openpyxl", skiprows=1)
    print(f"Colonne rilevate nel file Excel: {list(df.columns)}")

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
    print(f"\nErrore durante l'elaborazione: {e}")
    sys.exit(1)

import os
import sys
import json
import pandas as pd
from curl_cffi import requests

email = os.environ.get("FANTA_EMAIL")
password = os.environ.get("FANTA_PASSWORD")

if not email or not password:
    print("ERRORE: Le variabili d'ambiente FANTA_EMAIL o FANTA_PASSWORD non sono state configurate nei GitHub Secrets.")
    sys.exit(1)

session = requests.Session(impersonate="chrome120")

login_url = "https://www.fantacalcio.it/api/v1/User/login"
download_url = "https://www.fantacalcio.it/api/v1/Excel/prices/21/1"

try:
    print("1. Inizializzazione sessione e raccolta cookie base...")
    session.get("https://www.fantacalcio.it", timeout=20)

    print("2. Tentativo di autenticazione con credenziali protette...")
    login_payload = {
        "username": email,
        "password": password,
        "rememberMe": True
    }
    
    headers_auth = {
        "Referer": "https://www.fantacalcio.it/",
        "Origin": "https://www.fantacalcio.it",
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*"
    }

    res_login = session.post(login_url, json=login_payload, headers=headers_auth, timeout=25)
    print(f"   Esito Login HTTP: {res_login.status_code}")

    # Estrazione dell'eventuale Bearer Token se presente nella risposta JSON
    auth_token = None
    try:
        login_data = res_login.json()
        auth_token = login_data.get("token") or login_data.get("data", {}).get("token")
        if auth_token:
            print("   Token di autorizzazione Bearer rilevato.")
    except Exception:
        pass

    print("\n3. Download del listone Excel...")
    headers_download = {
        "Referer": "https://www.fantacalcio.it/quotazioni-fantacalcio",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }
    if auth_token:
        headers_download["Authorization"] = f"Bearer {auth_token}"

    res_file = session.get(download_url, headers=headers_download, timeout=30)
    print(f"   Status HTTP Download: {res_file.status_code}")
    print(f"   Dimensione scaricata: {len(res_file.content)} bytes")

    with open("quotazioni.xlsx", "wb") as f:
        f.write(res_file.content)

    primi_byte = res_file.content[:10]
    print(f"   Primi byte: {primi_byte}")

    if not primi_byte.startswith(b"PK\x03\x04"):
        print("\nERRORE: La risposta non è un file Excel valido.")
        try:
            print("Dettaglio risposta:", res_file.text[:400])
        except Exception:
            pass
        sys.exit(1)

    print("\nFile Excel autenticato con successo! Elaborazione dati in corso...")
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

    print(f"\nOperazione riuscita: generato 'dati_serie_a.json' con {len(players)} calciatori.")

except Exception as e:
    print(f"\nErrore imprevisto durante l'esecuzione: {e}")
    sys.exit(1)

import os
import sys
import json
import pandas as pd
from curl_cffi import requests

email = os.environ.get("FANTA_EMAIL")
password = os.environ.get("FANTA_PASSWORD")

if not email or not password:
    print("ERRORE: Credenziali mancanti nei Secrets.")
    sys.exit(1)

session = requests.Session(impersonate="chrome120")

login_url = "https://www.fantacalcio.it/api/v1/User/login"
download_url = "https://www.fantacalcio.it/api/v1/Excel/prices/21/1"

def pulisci_numero(valore, default=0.0):
    if pd.isnull(valore):
        return default
    if isinstance(valore, (int, float)):
        return float(valore)
    try:
        val_str = str(valore).strip().replace(",", ".")
        return float(val_str)
    except Exception:
        return default

try:
    print("1. Inizializzazione sessione...")
    session.get("https://www.fantacalcio.it", timeout=20)

    print("2. Login account...")
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

    auth_token = None
    try:
        data = res_login.json()
        auth_token = data.get("token") or data.get("data", {}).get("token")
    except Exception:
        pass

    print("3. Download file Excel...")
    headers_down = {
        "Referer": "https://www.fantacalcio.it/quotazioni-fantacalcio",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }
    if auth_token:
        headers_down["Authorization"] = f"Bearer {auth_token}"

    res_file = session.get(download_url, headers=headers_down, timeout=30)

    with open("quotazioni.xlsx", "wb") as f:
        f.write(res_file.content)

    # Legge l'Excel
    df = pd.read_excel("quotazioni.xlsx", engine="openpyxl", skiprows=1)

    # Rende tutti i nomi delle colonne in MAIUSCOLO e senza spazi vuoti
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"Colonne rilevate nel file: {list(df.columns)}")

    players = []
    for _, row in df.iterrows():
        try:
            nome = str(row.get("NOME", "")).strip()
            if not nome or nome.lower() == "nan":
                continue

            # Gestione dinamica dei possibili nomi colonna
            ruolo = str(row.get("R", row.get("RUOLO", ""))).strip()
            squadra = str(row.get("SQUADRA", "")).strip()
            
            # Quotazione (Qt.A o QUOTAZIONE)
            qt = int(pulisci_numero(row.get("QT.A", row.get("QT", row.get("QUOTAZIONE", 1))), default=1))
            
            # Statistiche: gestione FM, MV, presenze e bonus/malus
            fm = pulisci_numero(row.get("FM", row.get("FANTAMEDIA", 0.0)))
            mv = pulisci_numero(row.get("MV", row.get("MEDIAVOTO", row.get("MEDIA VOTO", 0.0))))
            
            # Presenze (PV o PG)
            pv = int(pulisci_numero(row.get("PV", row.get("PG", row.get("PRESENZE", 0))), default=0))
            
            # Gol, Assist, Malus
            gf = int(pulisci_numero(row.get("GF", row.get("GOL", 0)), default=0))
            ass = int(pulisci_numero(row.get("ASS", row.get("AS", row.get("ASSIST", 0))), default=0))
            amm = int(pulisci_numero(row.get("AMM", row.get("AMMONIZIONI", 0)), default=0))
            esp = int(pulisci_numero(row.get("ESP", row.get("ESPULSIONI", 0)), default=0))

            players.append({
                "id": int(row.get("ID", 0)) if pd.notnull(row.get("ID")) else 0,
                "ruolo": ruolo,
                "nome": nome,
                "squadra": squadra,
                "quotazione": qt,
                "fantaMedia": round(fm, 2),
                "mediaVoto": round(mv, 2),
                "presenze": pv,
                "gol": gf,
                "assist": ass,
                "ammonizioni": amm,
                "espulsioni": esp,
                "status": "disponibile"
            })
        except Exception:
            continue

    with open("dati_serie_a.json", "w", encoding="utf-8") as f:
        json.dump(players, f, ensure_ascii=False, indent=2)

    print(f"Completato con successo: {len(players)} giocatori elaborati con statistiche reali.")

except Exception as e:
    print(f"Errore durante l'elaborazione: {e}")
    sys.exit(1)

import os
import sys
import json
import pandas as pd
from curl_cffi import requests

email = os.environ.get("FANTA_EMAIL")
password = os.environ.get("FANTA_PASSWORD")

if not email or not password:
    print("ERRORE: Credenziali mancanti nei GitHub Secrets.")
    sys.exit(1)

session = requests.Session(impersonate="chrome120")

login_url = "https://www.fantacalcio.it/api/v1/User/login"
url_prices = "https://www.fantacalcio.it/api/v1/Excel/prices/21/1"
url_stats = "https://www.fantacalcio.it/api/v1/Excel/stats/21/1"

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
    print("1. Login a Fantacalcio.it...")
    session.get("https://www.fantacalcio.it", timeout=20)

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

    headers_down = {
        "Referer": "https://www.fantacalcio.it/quotazioni-fantacalcio",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }
    if auth_token:
        headers_down["Authorization"] = f"Bearer {auth_token}"

    # 1. DOWNLOAD QUOTAZIONI
    print("\n2. Download file Quotazioni (prices)...")
    res_prices = session.get(url_prices, headers=headers_down, timeout=30)
    with open("quotazioni.xlsx", "wb") as f:
        f.write(res_prices.content)

    df_prices = pd.read_excel("quotazioni.xlsx", sheet_name=0, engine="openpyxl", skiprows=1)
    df_prices.columns = [str(c).strip().upper() for c in df_prices.columns]
    print(f"Quotazioni caricate. Colonne: {list(df_prices.columns)}")

    # 2. DOWNLOAD STATISTICHE
    print("\n3. Download file Statistiche (stats)...")
    headers_stats = headers_down.copy()
    headers_stats["Referer"] = "https://www.fantacalcio.it/statistiche-fantacalcio"

    res_stats = session.get(url_stats, headers=headers_stats, timeout=30)
    with open("statistiche.xlsx", "wb") as f:
        f.write(res_stats.content)

    df_stats = pd.read_excel("statistiche.xlsx", sheet_name=0, engine="openpyxl", skiprows=1)
    df_stats.columns = [str(c).strip().upper() for c in df_stats.columns]
    print(f"Statistiche caricate. Colonne: {list(df_stats.columns)}")

    # 3. MAPPATURA STATISTICHE PER ID
    stats_by_id = {}
    for _, row in df_stats.iterrows():
        try:
            p_id = int(row.get("ID", 0)) if pd.notnull(row.get("ID")) else 0
            if p_id == 0:
                continue

            fm = pulisci_numero(row.get("FM", row.get("FANTAMEDIA", 0.0)))
            mv = pulisci_numero(row.get("MV", row.get("MEDIAVOTO", 0.0)))
            pv = int(pulisci_numero(row.get("PV", row.get("PG", 0))))
            gf = int(pulisci_numero(row.get("GF", row.get("GOL", 0))))
            ass = int(pulisci_numero(row.get("ASS", row.get("AS", 0))))
            amm = int(pulisci_numero(row.get("AMM", 0)))
            esp = int(pulisci_numero(row.get("ESP", 0)))

            stats_by_id[p_id] = {
                "fantaMedia": round(fm, 2),
                "mediaVoto": round(mv, 2),
                "presenze": pv,
                "gol": gf,
                "assist": ass,
                "ammonizioni": amm,
                "espulsioni": esp
            }
        except Exception:
            continue

    # 4. UNIONE CON QUOTAZIONI E CREAZIONE DATASET
    players = []
    for _, row in df_prices.iterrows():
        try:
            nome = str(row.get("NOME", "")).strip()
            if not nome or nome.lower() == "nan":
                continue

            p_id = int(row.get("ID", 0)) if pd.notnull(row.get("ID")) else 0
            ruolo = str(row.get("R", row.get("RUOLO", ""))).strip()
            squadra = str(row.get("SQUADRA", "")).strip()
            qt = int(pulisci_numero(row.get("QT.A", row.get("QT", 1)), default=1))

            st = stats_by_id.get(p_id, {
                "fantaMedia": 0.0,
                "mediaVoto": 0.0,
                "presenze": 0,
                "gol": 0,
                "assist": 0,
                "ammonizioni": 0,
                "espulsioni": 0
            })

            players.append({
                "id": p_id,
                "ruolo": ruolo,
                "nome": nome,
                "squadra": squadra,
                "quotazione": qt,
                "fantaMedia": st["fantaMedia"],
                "mediaVoto": st["mediaVoto"],
                "presenze": st["presenze"],
                "gol": st["gol"],
                "assist": st["assist"],
                "ammonizioni": st["ammonizioni"],
                "espulsioni": st["espulsioni"],
                "status": "disponibile"
            })
        except Exception:
            continue

    with open("dati_serie_a.json", "w", encoding="utf-8") as f:
        json.dump(players, f, ensure_ascii=False, indent=2)

    print(f"\nOperazione completata con successo: generati {len(players)} giocatori.")
    print("Esempio primo giocatore elaborato:")
    if players:
        print(json.dumps(players[0], indent=2, ensure_ascii=False))

except Exception as e:
    print(f"\nErrore durante l'elaborazione: {e}")
    sys.exit(1)

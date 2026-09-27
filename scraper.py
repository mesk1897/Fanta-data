import pandas as pd
import requests
import json

# URL dell'export ufficiale Fantacalcio
url = "https://www.fantacalcio.it/servizi/Excel/Quotazioni_Fantacalcio_Stagione_2026_27.xlsx"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

try:
    response = requests.get(url, headers=headers, timeout=30)
    with open("quotazioni.xlsx", "wb") as f:
        f.write(response.content)

    df = pd.read_excel("quotazioni.xlsx", skiprows=1)

    players = []
    for _, row in df.iterrows():
        try:
            nome = str(row.get("Nome", "")).strip()
            if not nome or nome == "nan":
                continue
            players.append({
                "id": int(row.get("Id", 0)),
                "ruolo": str(row.get("R", "")).strip(),
                "nome": nome,
                "squadra": str(row.get("Squadra", "")).strip(),
                "quotazione": int(row.get("Qt.A", 1)),
                "fantaMedia": float(row.get("Fm", 0) if pd.notnull(row.get("Fm")) else 0),
                "mediaVoto": float(row.get("Mv", 0) if pd.notnull(row.get("Mv")) else 0),
                "gol": int(row.get("Gf", 0) if pd.notnull(row.get("Gf")) else 0),
                "assist": int(row.get("Ass", 0) if pd.notnull(row.get("Ass")) else 0),
                "ammonizioni": int(row.get("Amm", 0) if pd.notnull(row.get("Amm")) else 0),
                "espulsioni": int(row.get("Esp", 0) if pd.notnull(row.get("Esp")) else 0),
                "status": "disponibile"
            })
        except Exception:
            continue

    with open("dati_serie_a.json", "w", encoding="utf-8") as f:
        json.dump(players, f, ensure_ascii=False, indent=2)
    print(f"Salvati con successo {len(players)} giocatori.")
except Exception as e:
    print(f"Errore durante l'aggiornamento: {e}")
    raise e

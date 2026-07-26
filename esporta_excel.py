import sys
import os
import requests
import pandas as pd
from docx import Document

# ==========================================
# 1. SISTEMA DI AGGIORNAMENTO AUTOMATICO
# ==========================================
VERSIONE_CORRENTE = "v1.0.0"

def controlla_aggiornamenti():
    try:
        url = "https://api.github.com/repos/ReneFerretti/serie-a-engine/releases/latest"
        response = requests.get(url, timeout=3)
        if response.status_code == 200:
            data = response.json()
            ultima_versione = data["tag_name"]
            
            if ultima_versione != VERSIONE_CORRENTE:
                print(f"\n[!] È disponibile un nuovo aggiornamento: {ultima_versione} (la tua versione è {VERSIONE_CORRENTE})")
                scelta = input("Vuoi scaricare l'aggiornamento ora? (s/n): ").strip().lower()
                if scelta == 's':
                    for asset in data.get("assets", []):
                        if asset["name"].endswith(".exe"):
                            download_url = asset["browser_download_url"]
                            print("Scaricamento del nuovo file in corso...")
                            
                            r = requests.get(download_url)
                            with open("esporta_excel_nuovo.exe", "wb") as f:
                                f.write(r.content)
                                
                            print("Aggiornamento completato!")
                            print("Troverai il nuovo file come 'esporta_excel_nuovo.exe'. Sostituisci il vecchio con questo.")
                            input("Premi un tasto per uscire...")
                            sys.exit()
    except Exception:
        pass

# Esegui il controllo all'avvio
controlla_aggiornamenti()


# ==========================================
# 2. LOGICA DEL TUO PROGRAMMA (EXCEL E WORD)
# ==========================================
print("Avvio del programma: generazione file in corso...")

try:
    # Esempio di creazione dati per Excel
    dati = {
        "Squadra": ["Inter", "Milan", "Juventus", "Napoli"],
        "Punti": [45, 42, 40, 38],
        "xGoal": [35.5, 32.1, 30.4, 29.8]
    }
    df = pd.DataFrame(dati)
    
    nome_file_excel = "classifica_serie_a.xlsx"
    df.to_excel(nome_file_excel, index=False)
    print(f"File Excel generato con successo: {nome_file_excel}")

    # Esempio di creazione documento Word
    doc = Document()
    doc.add_heading("Report Aggiornato Serie A", 0)
    doc.add_paragraph("Questo report è stato generato e aggiornato automaticamente dal programma.")
    
    for index, row in df.iterrows():
        doc.add_paragraph(f"{row['Squadra']} - Punti: {row['Punti']} (xG: {row['xGoal']})")
        
    nome_file_word = "report_serie_a.docx"
    doc.save(nome_file_word)
    print(f"File Word generato con successo: {nome_file_word}")

except Exception as e:
    print(f"Si è verificato un errore durante l'elaborazione: {e}")

print("\nOperazione completata con successo!")
input("Premi un tasto per chiudere il programma...")

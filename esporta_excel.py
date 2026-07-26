import os
import requests
import pandas as pd
import numpy as np
from scipy.stats import poisson
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import subprocess

# 1. CONFIGURAZIONE E STORICO
API_KEY = 'de85dd8016d04173bf8cbe9a86a1a617'
url_storico = "https://www.football-data.co.uk/mmz4281/2526/I1.csv"
headers = {'X-Auth-Token': API_KEY}

print("Step 1: Caricamento dati storici e calcolo forza squadre...")
df_storico = pd.read_csv(url_storico).dropna(subset=['HomeTeam', 'AwayTeam', 'FTHG', 'FTAG'])

squadre = pd.unique(df_storico[['HomeTeam', 'AwayTeam']].values.ravel('K'))
forza_attacco_casa, forza_difesa_casa = {}, {}
forza_attacco_trasferta, forza_difesa_trasferta = {}, {}

media_gol_casa_lg = df_storico['FTHG'].mean()
media_gol_trasferta_lg = df_storico['FTAG'].mean()

for sq in squadre:
    p_casa = df_storico[df_storico['HomeTeam'] == sq]
    p_trasferta = df_storico[df_storico['AwayTeam'] == sq]
    
    f_att_c = p_casa['FTHG'].mean() if len(p_casa) > 0 else media_gol_casa_lg
    f_dif_c = p_casa['FTAG'].mean() if len(p_casa) > 0 else media_gol_trasferta_lg
    f_att_t = p_trasferta['FTAG'].mean() if len(p_trasferta) > 0 else media_gol_trasferta_lg
    f_dif_t = p_trasferta['FTHG'].mean() if len(p_trasferta) > 0 else media_gol_casa_lg
    
    forza_attacco_casa[sq] = f_att_c / media_gol_casa_lg
    forza_difesa_casa[sq] = f_dif_c / media_gol_trasferta_lg
    forza_attacco_trasferta[sq] = f_att_t / media_gol_trasferta_lg
    forza_difesa_trasferta[sq] = f_dif_t / media_gol_casa_lg

mapping_nomi = {
    "FC Internazionale Milano": "Inter",
    "AC Milan": "Milan",
    "Juventus FC": "Juventus",
    "SSC Napoli": "Napoli",
    "AS Roma": "Roma",
    "SS Lazio": "Lazio",
    "Atalanta BC": "Atalanta",
    "ACF Fiorentina": "Fiorentina",
    "Bologna FC 1909": "Bologna",
    "Torino FC": "Torino",
    "Udinese Calcio": "Udinese",
    "Genoa CFC": "Genoa",
    "AC Monza": "Monza",
    "US Lecce": "Lecce",
    "Cagliari Calcio": "Cagliari",
    "Hellas Verona FC": "Verona",
    "Empoli FC": "Empoli",
    "Parma Calcio 1913": "Parma",
    "Venezia FC": "Venezia",
    "Como 1907": "Como"
}

# 2. SCARICAMENTO ROSE E ANALISI PROFONDITÀ SQUADRE
print("Step 2: Analisi delle rose e dei roster ufficiali...")
url_teams = "https://api.football-data.org/v4/competitions/2019/teams"
res_teams = requests.get(url_teams, headers=headers)
dati_rose = {}

if res_teams.status_code == 200:
    teams_data = res_teams.json().get('teams', [])
    for t in teams_data:
        raw_t_name = t['name']
        nome_pulito = mapping_nomi.get(raw_t_name, raw_t_name)
        squadra_giocatori = t.get('squad', [])
        
        num_giocatori = len(squadra_giocatori)
        attaccanti = sum(1 for p in squadra_giocatori if p.get('position') == 'Offence')
        centrocampisti = sum(1 for p in squadra_giocatori if p.get('position') == 'Midfield')
        difensori = sum(1 for p in squadra_giocatori if p.get('position') == 'Defence')
        portieri = sum(1 for p in squadra_giocatori if p.get('position') == 'Goalkeeper')
        
        dati_rose[nome_pulito] = {
            'totale_rosa': num_giocatori,
            'attaccanti': attaccanti,
            'centrocampisti': centrocampisti,
            'difensori': difensori,
            'portieri': portieri
        }

# 3. GESTIONE DINAMICA INFORTUNI E ASSENZE PESANTI
coeff_infortuni = {
    "Inter": 1.0, "Milan": 1.0, "Juventus": 0.95, "Napoli": 1.0, "Roma": 1.0,
    "Lazio": 1.0, "Atalanta": 0.97, "Fiorentina": 1.0, "Bologna": 1.0, "Torino": 1.0,
    "Udinese": 1.0, "Genoa": 1.0, "Monza": 1.0, "Lecce": 1.0, "Cagliari": 1.0,
    "Verona": 1.0, "Empoli": 1.0, "Parma": 1.0, "Venezia": 1.0, "Como": 1.0
}

# 4. SCARICAMENTO CLASSIFICA E STATISTICHE (Aggiornato con Casa e Trasferta)
print("Step 3: Connessione all'API per classifica generale, casa e trasferta...")
url_standings = "https://api.football-data.org/v4/competitions/2019/standings"
res_standings = requests.get(url_standings, headers=headers)

classifiche_dettaglio = []
forma_recentissima = {}
dati_squadre_api = {}

if res_standings.status_code == 200:
    data_standings = res_standings.json()
    standings = data_standings.get('standings', [])
    
    tot_table, home_table, away_table = [], [], []
    for table in standings:
        t_type = table.get('type')
        if t_type == 'TOTAL': tot_table = table.get('table', [])
        elif t_type == 'HOME': home_table = table.get('table', [])
        elif t_type == 'AWAY': away_table = table.get('table', [])
        
    home_dict = {mapping_nomi.get(row['team']['name'], row['team']['name']): row for row in home_table}
    away_dict = {mapping_nomi.get(row['team']['name'], row['team']['name']): row for row in away_table}
    
    for pos in tot_table:
        raw_name = pos['team']['name']
        squadra_pulita = mapping_nomi.get(raw_name, raw_name)
        
        form_str = pos.get('form', '') 
        punti_forma = 0
        if form_str:
            for char in form_str[-5:]:
                if char == 'W': punti_forma += 3
                elif char == 'D': punti_forma += 1
        coeff_forma = 1.0 + ((punti_forma - 7.5) / 30.0) if form_str else 1.0
        forma_recentissima[squadra_pulita] = max(0.8, min(1.2, coeff_forma))

        dati_squadre_api[squadra_pulita] = {
            'pos': pos.get('position', 0),
            'pt': pos.get('points', 0),
            'gf': pos.get('goalsFor', 0),
            'gs': pos.get('goalsAgainst', 0),
            'form': form_str if form_str else "N/D"
        }

        # Dati specifici Casa e Trasferta dai dizionari dedicati
        h_data = home_dict.get(squadra_pulita, {})
        a_data = away_dict.get(squadra_pulita, {})

        classifiche_dettaglio.append({
            'Pos': pos.get('position', 0),
            'Squadra': squadra_pulita,
            'Rosa (Num)': dati_rose.get(squadra_pulita, {}).get('totale_rosa', '-'),
            'Forma (Ultime)': form_str,
            'Pt': pos.get('points', 0),
            'G': pos.get('playedGames', 0),
            'V': pos.get('won', 0),
            'N': pos.get('draw', 0),
            'P': pos.get('lost', 0),
            'GF': pos.get('goalsFor', 0),
            'GS': pos.get('goalsAgainst', 0),
            'DR': pos.get('goalDifference', 0),
            # Nuove colonne Casa
            'Casa G': h_data.get('playedGames', 0),
            'Casa V': h_data.get('won', 0),
            'Casa N': h_data.get('draw', 0),
            'Casa P': h_data.get('lost', 0),
            'Casa GF': h_data.get('goalsFor', 0),
            'Casa GS': h_data.get('goalsAgainst', 0),
            # Nuove colonne Trasferta
            'Trasf G': a_data.get('playedGames', 0),
            'Trasf V': a_data.get('won', 0),
            'Trasf N': a_data.get('draw', 0),
            'Trasf P': a_data.get('lost', 0),
            'Trasf GF': a_data.get('goalsFor', 0),
            'Trasf GS': a_data.get('goalsAgainst', 0)
        })

df_classifica = pd.DataFrame(classifiche_dettaglio)

# 5. SCARICAMENTO PARTITE E COSTRUZIONE CRONACA GIORNALISTICA
url_api_all = "https://api.football-data.org/v4/competitions/2019/matches"
response_all = requests.get(url_api_all, headers=headers)
lista_risultati = []
report_lines = [] 

if response_all.status_code == 200:
    partite = response_all.json().get('matches', [])
    print(f"Trovate {len(partite)} partite totali. Generazione articoli in corso...")
    
    for match in partite:
        data = match['utcDate'][:10]
        giornata = match['matchday']
        stato = match['status']
        
        raw_casa = match['homeTeam']['name']
        raw_trasferta = match['awayTeam']['name']
        
        casa = mapping_nomi.get(raw_casa, raw_casa)
        trasferta = mapping_nomi.get(raw_trasferta, raw_trasferta)
        
        mg_tot = {f"{i}-{j}": 0 for i in range(1, 6) for j in range(i+1, 7)}
        mg_casa = {f"{i}-{j}": 0 for i in range(1, 4) for j in range(i+1, 5)}
        mg_tras = {f"{i}-{j}": 0 for i in range(1, 4) for j in range(i+1, 5)}
        
        if stato == 'FINISHED':
            gol_c_reali = match['score']['fullTime']['home']
            gol_t_reali = match['score']['fullTime']['away']
            risultato_reale = f"{gol_c_reali}-{gol_t_reali}"
            segno_reale = "1" if gol_c_reali > gol_t_reali else ("X" if gol_c_reali == gol_t_reali else "2")
            consiglio_testo = f"FINITA ({segno_reale} - {risultato_reale})"
            
            p1, px, p2 = (1.0, 0.0, 0.0) if segno_reale == "1" else (0.0, 1.0, 0.0) if segno_reale == "X" else (0.0, 0.0, 1.0)
            p_1x = p1 + px
            p_x2 = px + p2
            p_12 = p1 + p2
            p_over_25 = 1.0 if (gol_c_reali + gol_t_reali) > 2 else 0.0
            p_under_25 = 1.0 - p_over_25
            p_goal = 1.0 if (gol_c_reali > 0 and gol_t_reali > 0) else 0.0
            p_nogoal = 1.0 - p_goal
            
            top1_res, top1_prob = risultato_reale, 1.0
            top2_res, top2_prob = "-", 0.0
            top3_res, top3_prob = "-", 0.0
        else:
            f_casa_momentum = forma_recentissima.get(casa, 1.0) * coeff_infortuni.get(casa, 1.0)
            f_trasf_momentum = forma_recentissima.get(trasferta, 1.0) * coeff_infortuni.get(trasferta, 1.0)

            lambda_c = forza_attacco_casa.get(casa, 1.0) * forza_difesa_trasferta.get(trasferta, 1.0) * media_gol_casa_lg * f_casa_momentum
            lambda_t = forza_attacco_trasferta.get(trasferta, 1.0) * forza_difesa_casa.get(casa, 1.0) * media_gol_trasferta_lg * f_trasf_momentum
            
            prob_matrix = np.zeros((6, 6))
            for i in range(6):
                for j in range(6):
                    prob_matrix[i, j] = poisson.pmf(i, lambda_c) * poisson.pmf(j, lambda_t)
                    
            p1 = np.sum(np.tril(prob_matrix, -1))
            px = np.sum(np.diag(prob_matrix))
            p2 = np.sum(np.triu(prob_matrix, 1))
            
            p_1x = p1 + px
            p_x2 = px + p2
            p_12 = p1 + p2
            
            p_under_25 = 0
            p_goal = 0
            tutti_risultati_esatti = []
            
            for i in range(6):
                for j in range(6):
                    p_res = prob_matrix[i, j]
                    tot_gol = i + j
                    tutti_risultati_esatti.append((f"{i}-{j}", p_res))
                    
                    if tot_gol <= 2: p_under_25 += p_res
                    if i > 0 and j > 0: p_goal += p_res
                        
                    for mg in mg_tot.keys():
                        min_g, max_g = map(int, mg.split('-'))
                        if min_g <= tot_gol <= max_g: mg_tot[mg] += p_res
                    for mg in mg_casa.keys():
                        min_g, max_g = map(int, mg.split('-'))
                        if min_g <= i <= max_g: mg_casa[mg] += p_res
                    for mg in mg_tras.keys():
                        min_g, max_g = map(int, mg.split('-'))
                        if min_g <= j <= max_g: mg_tras[mg] += p_res
                        
            p_over_25 = 1.0 - p_under_25
            p_nogoal = 1.0 - p_goal
            
            tutti_risultati_esatti.sort(key=lambda x: x[1], reverse=True)
            top1_res, top1_prob = tutti_risultati_esatti[0]
            top2_res, top2_prob = tutti_risultati_esatti[1]
            top3_res, top3_prob = tutti_risultati_esatti[2]
            
            candidati = [
                ("1X", p_1x), ("X2", p_x2), ("12", p_12),
                ("1", p1), ("X", px), ("2", p2),
                ("Over 2.5", p_over_25), ("Under 2.5", p_under_25),
                ("Goal", p_goal), ("NoGoal", p_nogoal)
            ]
            for k, v in mg_tot.items(): candidati.append((f"Multigol {k}", v))
            for k, v in mg_casa.items(): candidati.append((f"MG Casa {k}", v))
            for k, v in mg_tras.items(): candidati.append((f"MG Trasf {k}", v))
            
            miglior_scelta = max(candidati, key=lambda x: x[1])
            consiglio_testo = f"{miglior_scelta[0]} ({round(miglior_scelta[1]*100, 1)}%)"
            
            inf_c = dati_squadre_api.get(casa, {})
            inf_t = dati_squadre_api.get(trasferta, {})
            rosa_c = dati_rose.get(casa, {})
            rosa_t = dati_rose.get(trasferta, {})
            
            report_lines.append(f"STADIO & MATCH: {casa} vs {trasferta} (Giornata {giornata} - {data})")
            
            incipit = f"L'attesa è finita. La lavagna tattica della {giornata}ª giornata ci offre un banco di prova affascinante tra i padroni di casa del {casa} e la compagine ospite del {trasferta}."
            
            nota_infortuni_c = "a pieno regime" if coeff_infortuni.get(casa, 1.0) >= 1.0 else "condizionata da alcune pesanti assenze nell'elenco dei convocati"
            nota_infortuni_t = "al completo" if coeff_infortuni.get(trasferta, 1.0) >= 1.0 else "costretta a ridisegnare l'assetto per via di defezioni importanti"
            
            contesto_classifica = (
                f"Analizzando il momento e la profondità delle rose (pari a {rosa_c.get('totale_rosa','?')} elementi per i padroni di casa e "
                f"{rosa_t.get('totale_rosa','?')} per gli ospiti), il {casa} occupa attualmente la {inf_c.get('pos','?')}ª posizione "
                f"con {inf_c.get('pt','?')} punti (score: {inf_c.get('form','N/D')}), arrivando a questa sfida {nota_infortuni_c}. "
                f"Dall'altra parte, il {trasferta} risponde stazionando al {inf_t.get('pos','?')}º posto con {inf_t.get('pt','?')} punti "
                f"e un'infermeria che si presenta {nota_infortuni_t}. Un duello scacchistico che pesa non poco sulle scelte dei tecnici."
            )
            
            lettura_tecnica = (
                f"Cosa dicono i modelli previsionali ricalcolati sulle condizioni effettive dei roster? L'attesa di reti stimata "
                f"si fissa su {round(lambda_c, 2)} gol per i padroni di casa e {round(lambda_t, 2)} per gli ospiti. "
                f"La bilancia delle probabilità pende verso il segno {miglior_scelta[0]} stimato intorno al {round(miglior_scelta[1]*100,1)}%, "
                f"tenendo conto del fattore campo e delle rotazioni a disposizione. "
                f"Nei radar dei risultati esatti, le chiavi di lettura principali premiano il {top1_res} ({round(top1_prob*100,1)}%) "
                f"e l'opzione secondaria del {top2_res} ({round(top2_prob*100,1)}%)."
            )
            
            report_lines.append(incipit)
            report_lines.append(contesto_classifica)
            report_lines.append(lettura_tecnica)
            report_lines.append("")

        riga = {
            'Giornata': giornata,
            'Data': data,
            'Stato': 'Giocata' if stato == 'FINISHED' else 'Da Giocare',
            'Squadra Casa': casa,
            'Squadra Trasferta': trasferta,
            'Consiglio / Risultato': consiglio_testo,
            'Prob 1': round(p1, 4), 'Quota 1': round(1/p1, 2) if p1 > 0 else 0,
            'Prob X': round(px, 4), 'Quota X': round(1/px, 2) if px > 0 else 0,
            'Prob 2': round(p2, 4), 'Quota 2': round(1/p2, 2) if p2 > 0 else 0,
            'Prob 1X': round(p_1x, 4), 'Quota 1X': round(1/p_1x, 2) if p_1x > 0 else 0,
            'Prob X2': round(p_x2, 4), 'Quota X2': round(1/p_x2, 2) if p_x2 > 0 else 0,
            'Prob 12': round(p_12, 4), 'Quota 12': round(1/p_12, 2) if p_12 > 0 else 0,
            'Prob Over 2.5': round(p_over_25, 4), 'Quota Over': round(1/p_over_25, 2) if p_over_25 > 0 else 0,
            'Prob Under 2.5': round(p_under_25, 4), 'Quota Under': round(1/p_under_25, 2) if p_under_25 > 0 else 0,
            'Prob Goal': round(p_goal, 4), 'Quota Goal': round(1/p_goal, 2) if p_goal > 0 else 0,
            'Prob NoGoal': round(p_nogoal, 4), 'Quota NoGoal': round(1/p_nogoal, 2) if p_nogoal > 0 else 0,
        }
        
        for k, v in mg_tot.items():
            riga[f'Prob MG {k}'] = round(v, 4)
            riga[f'Quota MG {k}'] = round(1/v, 2) if v > 0 else 0
        for k, v in mg_casa.items():
            riga[f'Prob MG Casa {k}'] = round(v, 4)
            riga[f'Quota MG Casa {k}'] = round(1/v, 2) if v > 0 else 0
        for k, v in mg_tras.items():
            riga[f'Prob MG Trasf {k}'] = round(v, 4)
            riga[f'Quota MG Trasf {k}'] = round(1/v, 2) if v > 0 else 0
            
        riga.update({
            '1° Risultato': top1_res,
            'Quota Esatto 1': round(1/top1_prob, 2) if top1_prob > 0 else 0,
            '2° Risultato': top2_res,
            'Quota Esatto 2': round(1/top2_prob, 2) if top2_prob > 0 else 0,
            '3° Risultato': top3_res,
            'Quota Esatto 3': round(1/top3_prob, 2) if top3_prob > 0 else 0,
        })
        
        lista_risultati.append(riga)

    df_output = pd.DataFrame(lista_risultati)
    scrivania = os.path.expanduser("~/Desktop")
    file_name = os.path.join(scrivania, "scanner_serie_a.xlsx")
    file_report = os.path.join(scrivania, "Analisi_e_Consigli_Serie_A.rtf")
    
    wb = openpyxl.Workbook()
    
    ws1 = wb.active
    ws1.title = "Calendario e Quote"
    ws1.views.sheetView[0].showGridLines = True
    headers1 = list(df_output.columns)
    ws1.append(headers1)
    
    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    
    for col in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for row in dataframe_to_rows(df_output, index=False, header=False):
        ws1.append(row)
        
    ws1.auto_filter.ref = ws1.dimensions
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    ws2 = wb.create_sheet(title="Classifica e Rose")
    ws2.views.sheetView[0].showGridLines = True
    headers2 = list(df_classifica.columns)
    ws2.append(headers2)
    
    for col in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for row in dataframe_to_rows(df_classifica, index=False, header=False):
        ws2.append(row)
        
    ws2.auto_filter.ref = ws2.dimensions
    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(file_name)

    rtf_content = r"{\rtf1\ansi\ansicpg1252\deff0\deflang1040" + "\n"
    rtf_content += r"{\fonttbl{\f0\fnil\fcharset0 Arial;}}" + "\n"
    rtf_content += r"{\colortbl ;\red15\green23\blue42;\red180\red30\red30;\red70\red70\red70;}" + "\n"
    rtf_content += r"\viewkind4\uc1\pard\sa200\sl276\slmult1\f0\fs22 " + "\n"
    
    rtf_content += r"\b\fs36\cf1 IL FOCUS DELLA GIORNATA - ROSE & PRONOSTICI SERIE A\b0\fs22\cf0\par" + "\n"
    rtf_content += r"\i Editoriale tattico integrato con l'analisi della profondità delle rose e il peso degli infortuni sulle previsioni di Poisson.\i0\par\line" + "\n"
    
    for line in report_lines:
        safe_line = line.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')
        if "STADIO & MATCH:" in line:
            rtf_content += r"\b\fs26\cf2\sa120\sb240 " + safe_line + r"\b0\fs22\cf0\par" + "\n"
        elif line == "":
            rtf_content += r"\pard\sa300\par" + "\n"
        else:
            rtf_content += safe_line + r" " + "\n"
            
    rtf_content += r"}"

    with open(file_report, "w", encoding="utf-8") as f:
        f.write(rtf_content)

    print(f"\nFatto! File aggiornati con statistiche casa/trasferta nel foglio classifica.")
    
    subprocess.run(["open", file_name])
    subprocess.run(["open", file_report])
else:
    print(f"Errore API Classifiche/Partite: {response_all.status_code}")
    
#!/usr/bin/env python3
"""
Script per creare il Portfolio Tracker Excel
Importabile in Google Sheets per utilizzare GOOGLEFINANCE
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def create_portfolio_tracker():
    wb = Workbook()

    # Rimuovi il foglio default
    wb.remove(wb.active)

    # Stili comuni
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 1. DASHBOARD
    ws_dashboard = wb.create_sheet("Dashboard")
    dashboard_headers = [
        ["RIEPILOGO PATRIMONIO", ""],
        ["", ""],
        ["Categoria", "Valore (€)"],
    ]

    for row_idx, row_data in enumerate(dashboard_headers, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws_dashboard.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14)
            elif row_idx == 3:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")

    categories = [
        "Azioni",
        "ETF",
        "Bond",
        "Fondi Comuni",
        "Conti & Liquidità",
        "TOTALE ATTIVO",
        "",
        "Mutuo & Debiti",
        "TOTALE PASSIVO",
        "",
        "PATRIMONIO NETTO"
    ]

    for idx, cat in enumerate(categories, 4):
        ws_dashboard.cell(row=idx, column=1, value=cat)
        if cat in ["TOTALE ATTIVO", "TOTALE PASSIVO", "PATRIMONIO NETTO"]:
            ws_dashboard.cell(row=idx, column=1).font = Font(bold=True)
            ws_dashboard.cell(row=idx, column=2).font = Font(bold=True)

    # Formule dashboard (da aggiornare dopo aver creato gli altri fogli)
    ws_dashboard.cell(row=4, column=2, value="=SUMIF(Azioni!A:A,\"<>\",Azioni!H:H)")
    ws_dashboard.cell(row=5, column=2, value="=SUMIF(ETF!A:A,\"<>\",ETF!H:H)")
    ws_dashboard.cell(row=6, column=2, value="=SUMIF(Bond!A:A,\"<>\",Bond!G:G)")
    ws_dashboard.cell(row=7, column=2, value="=SUMIF('Fondi Comuni'!A:A,\"<>\",'Fondi Comuni'!F:F)")
    ws_dashboard.cell(row=8, column=2, value="=SUMIF('Conti & Liquidità'!A:A,\"<>\",'Conti & Liquidità'!C:C)")
    ws_dashboard.cell(row=9, column=2, value="=SUM(B4:B8)")
    ws_dashboard.cell(row=11, column=2, value="=SUMIF('Mutuo & Debiti'!A:A,\"<>\",'Mutuo & Debiti'!C:C)")
    ws_dashboard.cell(row=12, column=2, value="=B11")
    ws_dashboard.cell(row=14, column=2, value="=B9-B12")

    ws_dashboard.column_dimensions['A'].width = 25
    ws_dashboard.column_dimensions['B'].width = 18

    # 2. AZIONI
    ws_azioni = wb.create_sheet("Azioni")
    azioni_headers = ["Ticker", "Nome", "Quantità", "Prezzo Carico (€)", "Prezzo Corrente (€)",
                     "Valore Carico (€)", "Valore Corrente (€)", "P&L (€)", "P&L (%)",
                     "Dividendi Ricevuti (€)", "Data Acquisto", "Note"]

    for col_idx, header in enumerate(azioni_headers, 1):
        cell = ws_azioni.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    # Esempio di riga con formule (riga 2)
    ws_azioni.cell(row=2, column=5, value='=IFERROR(GOOGLEFINANCE(A2,"price"),"")')
    ws_azioni.cell(row=2, column=6, value='=IF(C2<>"",C2*D2,"")')
    ws_azioni.cell(row=2, column=7, value='=IF(C2<>"",C2*E2,"")')
    ws_azioni.cell(row=2, column=8, value='=IF(G2<>"",G2-F2,"")')
    ws_azioni.cell(row=2, column=9, value='=IF(F2<>0,H2/F2*100,"")')

    # Larghezza colonne
    widths_azioni = [12, 20, 12, 18, 18, 18, 18, 15, 12, 20, 15, 30]
    for idx, width in enumerate(widths_azioni, 1):
        ws_azioni.column_dimensions[get_column_letter(idx)].width = width

    # 3. ETF
    ws_etf = wb.create_sheet("ETF")
    etf_headers = ["Ticker", "Nome", "Quantità", "Prezzo Carico (€)", "Prezzo Corrente (€)",
                   "Valore Carico (€)", "Valore Corrente (€)", "P&L (€)", "P&L (%)",
                   "Dividendi Ricevuti (€)", "Data Acquisto", "Note"]

    for col_idx, header in enumerate(etf_headers, 1):
        cell = ws_etf.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_etf.cell(row=2, column=5, value='=IFERROR(GOOGLEFINANCE(A2,"price"),"")')
    ws_etf.cell(row=2, column=6, value='=IF(C2<>"",C2*D2,"")')
    ws_etf.cell(row=2, column=7, value='=IF(C2<>"",C2*E2,"")')
    ws_etf.cell(row=2, column=8, value='=IF(G2<>"",G2-F2,"")')
    ws_etf.cell(row=2, column=9, value='=IF(F2<>0,H2/F2*100,"")')

    for idx, width in enumerate(widths_azioni, 1):
        ws_etf.column_dimensions[get_column_letter(idx)].width = width

    # 4. BOND
    ws_bond = wb.create_sheet("Bond")
    bond_headers = ["Nome/ISIN", "Valore Nominale (€)", "Prezzo Acquisto (%)", "Data Acquisto",
                   "Scadenza", "Cedola (%)", "Frequenza Cedole", "Valore Corrente (€)",
                   "Rendimento YTM (%)", "Cedole Incassate (€)", "Note"]

    for col_idx, header in enumerate(bond_headers, 1):
        cell = ws_bond.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_bond.cell(row=2, column=8, value='=IF(B2<>"",B2*(C2/100),"")')

    widths_bond = [20, 18, 18, 15, 15, 12, 18, 18, 18, 20, 30]
    for idx, width in enumerate(widths_bond, 1):
        ws_bond.column_dimensions[get_column_letter(idx)].width = width

    # 5. FONDI COMUNI
    ws_fondi = wb.create_sheet("Fondi Comuni")
    fondi_headers = ["Nome Fondo", "ISIN", "Quote", "Prezzo Carico (€)", "Prezzo Corrente (€)",
                    "Valore Corrente (€)", "P&L (€)", "P&L (%)", "Data Acquisto", "Note"]

    for col_idx, header in enumerate(fondi_headers, 1):
        cell = ws_fondi.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_fondi.cell(row=2, column=6, value='=IF(C2<>"",C2*E2,"")')
    ws_fondi.cell(row=2, column=7, value='=IF(F2<>"",F2-(C2*D2),"")')
    ws_fondi.cell(row=2, column=8, value='=IF(C2*D2<>0,G2/(C2*D2)*100,"")')

    widths_fondi = [25, 20, 10, 18, 18, 18, 15, 12, 15, 30]
    for idx, width in enumerate(widths_fondi, 1):
        ws_fondi.column_dimensions[get_column_letter(idx)].width = width

    # 6. CONTI & LIQUIDITÀ
    ws_conti = wb.create_sheet("Conti & Liquidità")
    conti_headers = ["Nome Conto/Banca", "Tipo", "Saldo (€)", "Tasso Interesse (%)",
                    "Scadenza Vincolo", "Interessi Maturati (€)", "Note"]

    for col_idx, header in enumerate(conti_headers, 1):
        cell = ws_conti.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    # Aggiungi tipo conto come esempio
    ws_conti.cell(row=2, column=2, value="Conto Corrente")
    ws_conti.cell(row=3, column=2, value="Deposito Libero")
    ws_conti.cell(row=4, column=2, value="Deposito Vincolato")

    widths_conti = [25, 20, 15, 18, 18, 20, 30]
    for idx, width in enumerate(widths_conti, 1):
        ws_conti.column_dimensions[get_column_letter(idx)].width = width

    # 7. MUTUO & DEBITI
    ws_mutuo = wb.create_sheet("Mutuo & Debiti")
    mutuo_headers = ["Tipo Debito", "Importo Iniziale (€)", "Debito Residuo (€)",
                    "Tasso Interesse (%)", "Rata Mensile (€)", "Data Inizio",
                    "Scadenza", "Note"]

    for col_idx, header in enumerate(mutuo_headers, 1):
        cell = ws_mutuo.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    widths_mutuo = [20, 20, 20, 18, 18, 15, 15, 30]
    for idx, width in enumerate(widths_mutuo, 1):
        ws_mutuo.column_dimensions[get_column_letter(idx)].width = width

    # 8. STIPENDI
    ws_stipendi = wb.create_sheet("Stipendi")
    stipendi_headers = ["Anno", "Mese", "Stipendio Lordo (€)", "Stipendio Netto (€)",
                       "Bonus/Extra (€)", "Totale Netto (€)", "Note"]

    for col_idx, header in enumerate(stipendi_headers, 1):
        cell = ws_stipendi.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_stipendi.cell(row=2, column=6, value='=IF(D2<>"",D2+E2,"")')

    # Aggiungi mesi come esempio
    current_year = datetime.now().year
    months = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
             "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

    for idx, month in enumerate(months, 2):
        ws_stipendi.cell(row=idx, column=1, value=current_year)
        ws_stipendi.cell(row=idx, column=2, value=month)

    widths_stipendi = [10, 15, 20, 20, 18, 18, 30]
    for idx, width in enumerate(widths_stipendi, 1):
        ws_stipendi.column_dimensions[get_column_letter(idx)].width = width

    # 9. SPESE MENSILI
    ws_spese = wb.create_sheet("Spese Mensili")
    spese_headers = ["Anno", "Mese", "Affitto/Mutuo (€)", "Bollette (€)", "Spesa (€)",
                    "Trasporti (€)", "Assicurazioni (€)", "Abbonamenti (€)",
                    "Svago (€)", "Altro (€)", "TOTALE (€)", "Note"]

    for col_idx, header in enumerate(spese_headers, 1):
        cell = ws_spese.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_spese.cell(row=2, column=11, value='=SUM(C2:J2)')

    for idx, month in enumerate(months, 2):
        ws_spese.cell(row=idx, column=1, value=current_year)
        ws_spese.cell(row=idx, column=2, value=month)

    widths_spese = [10, 15, 15, 12, 12, 12, 15, 15, 12, 12, 15, 30]
    for idx, width in enumerate(widths_spese, 1):
        ws_spese.column_dimensions[get_column_letter(idx)].width = width

    # 10. TASSE
    ws_tasse = wb.create_sheet("Tasse")
    tasse_headers = ["Anno Fiscale", "Tipo", "Descrizione", "Base Imponibile (€)",
                    "Aliquota (%)", "Importo Pagato (€)", "Data Pagamento", "Note"]

    for col_idx, header in enumerate(tasse_headers, 1):
        cell = ws_tasse.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    # Esempi di tipi di tasse
    ws_tasse.cell(row=2, column=2, value="Capital Gain")
    ws_tasse.cell(row=3, column=2, value="Dividendi")
    ws_tasse.cell(row=4, column=2, value="Interessi")

    ws_tasse.cell(row=2, column=6, value='=IF(D2<>"",D2*(E2/100),"")')

    widths_tasse = [15, 18, 25, 20, 15, 20, 18, 30]
    for idx, width in enumerate(widths_tasse, 1):
        ws_tasse.column_dimensions[get_column_letter(idx)].width = width

    # 11. STORICO PATRIMONIO
    ws_storico = wb.create_sheet("Storico Patrimonio")
    storico_headers = ["Data", "Azioni", "ETF", "Bond", "Fondi", "Liquidità",
                      "Totale Attivo", "Debiti", "Patrimonio Netto", "Var. % Mensile"]

    for col_idx, header in enumerate(storico_headers, 1):
        cell = ws_storico.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    ws_storico.cell(row=2, column=7, value='=SUM(B2:F2)')
    ws_storico.cell(row=2, column=9, value='=G2-H2')
    ws_storico.cell(row=2, column=10, value='=IF(I1<>0,(I2-I1)/I1*100,"")')

    widths_storico = [12, 15, 15, 15, 15, 15, 18, 15, 20, 18]
    for idx, width in enumerate(widths_storico, 1):
        ws_storico.column_dimensions[get_column_letter(idx)].width = width

    # Salva il file
    output_file = "Portfolio_Tracker.xlsx"
    wb.save(output_file)
    print(f"✅ File '{output_file}' creato con successo!")
    print(f"\n📊 Fogli creati:")
    print(f"  1. Dashboard - Riepilogo patrimonio")
    print(f"  2. Azioni - Con formula GOOGLEFINANCE per prezzi automatici")
    print(f"  3. ETF - Con formula GOOGLEFINANCE per prezzi automatici")
    print(f"  4. Bond - Obbligazioni e titoli di stato")
    print(f"  5. Fondi Comuni - Fondi di investimento")
    print(f"  6. Conti & Liquidità - Cash e depositi")
    print(f"  7. Mutuo & Debiti - Gestione debiti")
    print(f"  8. Stipendi - Storico mensile")
    print(f"  9. Spese Mensili - Tracking spese (opzionale)")
    print(f" 10. Tasse - Capital gain e imposte")
    print(f" 11. Storico Patrimonio - Performance nel tempo")
    print(f"\n📥 Importa il file in Google Sheets per attivare le formule GOOGLEFINANCE!")

if __name__ == "__main__":
    create_portfolio_tracker()

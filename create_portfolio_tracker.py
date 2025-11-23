#!/usr/bin/env python3
"""
Script per creare il Portfolio Tracker Excel
Importabile in Google Sheets per utilizzare GOOGLEFINANCE
Versione 2.0 - Design migliorato con foglio Investimenti unificato
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def create_portfolio_tracker():
    wb = Workbook()
    wb.remove(wb.active)

    # Palette colori moderna
    COLOR_PRIMARY = "1E3A8A"      # Blu scuro
    COLOR_SECONDARY = "3B82F6"    # Blu
    COLOR_SUCCESS = "10B981"      # Verde
    COLOR_WARNING = "F59E0B"      # Arancione
    COLOR_DANGER = "EF4444"       # Rosso
    COLOR_LIGHT = "F3F4F6"        # Grigio chiaro
    COLOR_HEADER = "1F2937"       # Grigio scuro
    COLOR_ACCENT = "8B5CF6"       # Viola

    # Stili
    title_font = Font(bold=True, size=24, color="FFFFFF")
    title_fill = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type="solid")

    subtitle_font = Font(bold=True, size=14, color=COLOR_PRIMARY)

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")

    kpi_label_font = Font(bold=True, size=11, color=COLOR_HEADER)
    kpi_value_font = Font(bold=True, size=18, color=COLOR_PRIMARY)
    kpi_fill = PatternFill(start_color=COLOR_LIGHT, end_color=COLOR_LIGHT, fill_type="solid")

    border_thin = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    border_thick = Border(
        left=Side(style='medium', color=COLOR_PRIMARY),
        right=Side(style='medium', color=COLOR_PRIMARY),
        top=Side(style='medium', color=COLOR_PRIMARY),
        bottom=Side(style='medium', color=COLOR_PRIMARY)
    )

    # ========== DASHBOARD ==========
    ws_dash = wb.create_sheet("Dashboard")

    # Titolo principale
    ws_dash.merge_cells('A1:F1')
    title_cell = ws_dash['A1']
    title_cell.value = "PORTFOLIO TRACKER"
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[1].height = 40

    # Data ultimo aggiornamento
    ws_dash.merge_cells('A2:F2')
    ws_dash['A2'].value = f"Ultimo aggiornamento: {datetime.now().strftime('%d/%m/%Y')}"
    ws_dash['A2'].font = Font(size=10, italic=True, color="6B7280")
    ws_dash['A2'].alignment = Alignment(horizontal="center")

    # Riga vuota
    ws_dash.row_dimensions[3].height = 10

    # KPI PRINCIPALI
    ws_dash['A4'].value = "PATRIMONIO NETTO"
    ws_dash['A4'].font = kpi_label_font
    ws_dash['A4'].alignment = Alignment(horizontal="center")

    ws_dash.merge_cells('A5:B5')
    ws_dash['A5'].value = "=B15-B17"  # Formula patrimonio netto
    ws_dash['A5'].font = kpi_value_font
    ws_dash['A5'].fill = kpi_fill
    ws_dash['A5'].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash['A5'].number_format = '€#,##0.00'
    ws_dash['A5'].border = border_thick
    ws_dash.row_dimensions[5].height = 35

    # KPI Totale Attivo
    ws_dash['D4'].value = "TOTALE ATTIVO"
    ws_dash['D4'].font = Font(bold=True, size=10, color=COLOR_SUCCESS)
    ws_dash['D4'].alignment = Alignment(horizontal="center")

    ws_dash.merge_cells('D5:E5')
    ws_dash['D5'].value = "=B15"
    ws_dash['D5'].font = Font(bold=True, size=14, color=COLOR_SUCCESS)
    ws_dash['D5'].fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    ws_dash['D5'].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash['D5'].number_format = '€#,##0.00'
    ws_dash['D5'].border = border_thin

    # KPI Totale Passivo
    ws_dash['F4'].value = "TOTALE PASSIVO"
    ws_dash['F4'].font = Font(bold=True, size=10, color=COLOR_DANGER)
    ws_dash['F4'].alignment = Alignment(horizontal="center")

    ws_dash['F5'].value = "=B17"
    ws_dash['F5'].font = Font(bold=True, size=14, color=COLOR_DANGER)
    ws_dash['F5'].fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    ws_dash['F5'].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash['F5'].number_format = '€#,##0.00'
    ws_dash['F5'].border = border_thin

    # Riga vuota
    ws_dash.row_dimensions[6].height = 15

    # Sezione COMPOSIZIONE PORTAFOGLIO
    ws_dash.merge_cells('A7:B7')
    ws_dash['A7'].value = "COMPOSIZIONE PORTAFOGLIO"
    ws_dash['A7'].font = subtitle_font
    ws_dash['A7'].alignment = Alignment(horizontal="left", vertical="center")

    # Nota per grafico a torta
    ws_dash.merge_cells('D7:F7')
    ws_dash['D7'].value = "📊 INSERISCI QUI IL GRAFICO A TORTA"
    ws_dash['D7'].font = Font(bold=True, size=11, color=COLOR_ACCENT, italic=True)
    ws_dash['D7'].fill = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")
    ws_dash['D7'].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash['D7'].border = border_thin

    # Headers tabella composizione
    ws_dash['A8'].value = "Categoria"
    ws_dash['A8'].font = header_font
    ws_dash['A8'].fill = header_fill
    ws_dash['A8'].alignment = Alignment(horizontal="center")
    ws_dash['A8'].border = border_thin

    ws_dash['B8'].value = "Valore (€)"
    ws_dash['B8'].font = header_font
    ws_dash['B8'].fill = header_fill
    ws_dash['B8'].alignment = Alignment(horizontal="center")
    ws_dash['B8'].border = border_thin

    ws_dash['C8'].value = "% Portafoglio"
    ws_dash['C8'].font = header_font
    ws_dash['C8'].fill = header_fill
    ws_dash['C8'].alignment = Alignment(horizontal="center")
    ws_dash['C8'].border = border_thin

    # Placeholder per grafico (merge celle D8:F20)
    ws_dash.merge_cells('D8:F20')
    cell_chart = ws_dash['D8']
    cell_chart.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    cell_chart.border = border_thin
    cell_chart.alignment = Alignment(horizontal="center", vertical="center")
    cell_chart.value = "Seleziona i dati A9:B14\ne inserisci un grafico a torta\n\n(Inserisci → Grafico → Torta)"
    cell_chart.font = Font(size=11, color="9CA3AF", italic=True)

    # Dati composizione
    categories_data = [
        ("Investimenti", "=SUMIF(Investimenti!B:B,\"<>\",Investimenti!I:I)", "=IF($B$15>0,B9/$B$15*100,0)"),
        ("Fondi Comuni", "=SUMIF('Fondi Comuni'!A:A,\"<>\",'Fondi Comuni'!F:F)", "=IF($B$15>0,B10/$B$15*100,0)"),
        ("Conti & Liquidità", "=SUMIF('Conti & Liquidità'!A:A,\"<>\",'Conti & Liquidità'!C:C)", "=IF($B$15>0,B11/$B$15*100,0)"),
    ]

    row = 9
    for cat, formula_val, formula_pct in categories_data:
        ws_dash.cell(row=row, column=1, value=cat)
        ws_dash.cell(row=row, column=1).font = Font(size=11)
        ws_dash.cell(row=row, column=1).border = border_thin
        ws_dash.cell(row=row, column=1).alignment = Alignment(horizontal="left")

        ws_dash.cell(row=row, column=2, value=formula_val)
        ws_dash.cell(row=row, column=2).number_format = '€#,##0.00'
        ws_dash.cell(row=row, column=2).border = border_thin
        ws_dash.cell(row=row, column=2).alignment = Alignment(horizontal="right")

        ws_dash.cell(row=row, column=3, value=formula_pct)
        ws_dash.cell(row=row, column=3).number_format = '0.00"%"'
        ws_dash.cell(row=row, column=3).border = border_thin
        ws_dash.cell(row=row, column=3).alignment = Alignment(horizontal="center")
        row += 1

    # Riga vuota
    row += 1

    # Subtotale investimenti per tipo
    ws_dash.merge_cells(f'A{row}:C{row}')
    ws_dash.cell(row=row, column=1, value="DETTAGLIO INVESTIMENTI")
    ws_dash.cell(row=row, column=1).font = Font(bold=True, size=10, color=COLOR_SECONDARY)
    ws_dash.cell(row=row, column=1).fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws_dash.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws_dash.cell(row=row, column=1).border = border_thin
    row += 1

    inv_details = [
        ("  • Azioni", '=SUMIFS(Investimenti!I:I,Investimenti!A:A,"Azione")'),
        ("  • ETF", '=SUMIFS(Investimenti!I:I,Investimenti!A:A,"ETF")'),
        ("  • Bond", '=SUMIFS(Investimenti!I:I,Investimenti!A:A,"Bond")'),
    ]

    for label, formula in inv_details:
        ws_dash.cell(row=row, column=1, value=label)
        ws_dash.cell(row=row, column=1).font = Font(size=10, color="4B5563")
        ws_dash.cell(row=row, column=1).border = border_thin

        ws_dash.merge_cells(f'B{row}:C{row}')
        ws_dash.cell(row=row, column=2, value=formula)
        ws_dash.cell(row=row, column=2).number_format = '€#,##0.00'
        ws_dash.cell(row=row, column=2).font = Font(size=10)
        ws_dash.cell(row=row, column=2).border = border_thin
        ws_dash.cell(row=row, column=2).alignment = Alignment(horizontal="right")
        row += 1

    # TOTALE ATTIVO
    ws_dash.cell(row=row, column=1, value="TOTALE ATTIVO")
    ws_dash.cell(row=row, column=1).font = Font(bold=True, size=11, color=COLOR_SUCCESS)
    ws_dash.cell(row=row, column=1).fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    ws_dash.cell(row=row, column=1).border = border_thick

    ws_dash.merge_cells(f'B{row}:C{row}')
    ws_dash.cell(row=row, column=2, value="=SUM(B9:B11)")
    ws_dash.cell(row=row, column=2).number_format = '€#,##0.00'
    ws_dash.cell(row=row, column=2).font = Font(bold=True, size=11, color=COLOR_SUCCESS)
    ws_dash.cell(row=row, column=2).fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    ws_dash.cell(row=row, column=2).border = border_thick
    ws_dash.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    row += 1

    # Riga vuota
    row += 1

    # TOTALE PASSIVO
    ws_dash.cell(row=row, column=1, value="Mutuo & Debiti")
    ws_dash.cell(row=row, column=1).font = Font(size=11)
    ws_dash.cell(row=row, column=1).border = border_thin

    ws_dash.merge_cells(f'B{row}:C{row}')
    ws_dash.cell(row=row, column=2, value="=SUMIF('Mutuo & Debiti'!A:A,\"<>\",'Mutuo & Debiti'!C:C)")
    ws_dash.cell(row=row, column=2).number_format = '€#,##0.00'
    ws_dash.cell(row=row, column=2).font = Font(size=11, color=COLOR_DANGER)
    ws_dash.cell(row=row, column=2).border = border_thin
    ws_dash.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    row += 1

    # Riga vuota
    row += 1

    # PATRIMONIO NETTO FINALE
    ws_dash.cell(row=row, column=1, value="PATRIMONIO NETTO")
    ws_dash.cell(row=row, column=1).font = Font(bold=True, size=12, color="FFFFFF")
    ws_dash.cell(row=row, column=1).fill = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type="solid")
    ws_dash.cell(row=row, column=1).border = border_thick

    ws_dash.merge_cells(f'B{row}:C{row}')
    ws_dash.cell(row=row, column=2, value="=B15-B17")
    ws_dash.cell(row=row, column=2).number_format = '€#,##0.00'
    ws_dash.cell(row=row, column=2).font = Font(bold=True, size=12, color="FFFFFF")
    ws_dash.cell(row=row, column=2).fill = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type="solid")
    ws_dash.cell(row=row, column=2).border = border_thick
    ws_dash.cell(row=row, column=2).alignment = Alignment(horizontal="right")

    # Larghezze colonne Dashboard
    ws_dash.column_dimensions['A'].width = 22
    ws_dash.column_dimensions['B'].width = 16
    ws_dash.column_dimensions['C'].width = 16
    ws_dash.column_dimensions['D'].width = 16
    ws_dash.column_dimensions['E'].width = 16
    ws_dash.column_dimensions['F'].width = 16

    # ========== INVESTIMENTI (unificato) ==========
    ws_inv = wb.create_sheet("Investimenti")

    inv_headers = [
        "Tipo", "Ticker/ISIN", "Nome", "Quantità/Nominale",
        "Prezzo Carico (€)", "Prezzo Corrente (€)", "Valore Carico (€)",
        "Valore Corrente (€)", "P&L (€)", "P&L (%)",
        "Dividendi/Cedole (€)", "Data Acquisto", "Scadenza", "Note"
    ]

    for col_idx, header in enumerate(inv_headers, 1):
        cell = ws_inv.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    # Formule esempio riga 2
    ws_inv.cell(row=2, column=6, value='=IF(A2="Bond","",IFERROR(GOOGLEFINANCE(B2,"price"),""))')
    ws_inv.cell(row=2, column=7, value='=IF(D2<>"",D2*E2,"")')
    ws_inv.cell(row=2, column=8, value='=IF(AND(D2<>"",F2<>""),D2*F2,"")')
    ws_inv.cell(row=2, column=9, value='=IF(H2<>"",H2-G2,"")')
    ws_inv.cell(row=2, column=10, value='=IF(G2<>0,I2/G2*100,"")')

    # Aggiungi alcuni esempi
    ws_inv.cell(row=2, column=1, value="Azione")
    ws_inv.cell(row=3, column=1, value="ETF")
    ws_inv.cell(row=4, column=1, value="Bond")

    # Larghezze colonne
    widths_inv = [12, 15, 25, 16, 16, 16, 16, 16, 14, 11, 18, 14, 14, 30]
    for idx, width in enumerate(widths_inv, 1):
        ws_inv.column_dimensions[get_column_letter(idx)].width = width

    # Colora le celle Tipo in base al valore
    for row_num in range(2, 5):
        cell = ws_inv.cell(row=row_num, column=1)
        if cell.value == "Azione":
            cell.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
            cell.font = Font(bold=True, color=COLOR_SECONDARY)
        elif cell.value == "ETF":
            cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
            cell.font = Font(bold=True, color=COLOR_SUCCESS)
        elif cell.value == "Bond":
            cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            cell.font = Font(bold=True, color=COLOR_WARNING)

    # ========== FONDI COMUNI ==========
    ws_fondi = wb.create_sheet("Fondi Comuni")
    fondi_headers = ["Nome Fondo", "ISIN", "Quote", "Prezzo Carico (€)", "Prezzo Corrente (€)",
                    "Valore Corrente (€)", "P&L (€)", "P&L (%)", "Data Acquisto", "Note"]

    for col_idx, header in enumerate(fondi_headers, 1):
        cell = ws_fondi.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_fondi.cell(row=2, column=6, value='=IF(C2<>"",C2*E2,"")')
    ws_fondi.cell(row=2, column=7, value='=IF(F2<>"",F2-(C2*D2),"")')
    ws_fondi.cell(row=2, column=8, value='=IF(C2*D2<>0,G2/(C2*D2)*100,"")')

    widths_fondi = [25, 20, 10, 18, 18, 18, 15, 12, 15, 30]
    for idx, width in enumerate(widths_fondi, 1):
        ws_fondi.column_dimensions[get_column_letter(idx)].width = width

    # ========== CONTI & LIQUIDITÀ ==========
    ws_conti = wb.create_sheet("Conti & Liquidità")
    conti_headers = ["Nome Conto/Banca", "Tipo", "Saldo (€)", "Tasso Interesse (%)",
                    "Scadenza Vincolo", "Interessi Maturati (€)", "Note"]

    for col_idx, header in enumerate(conti_headers, 1):
        cell = ws_conti.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_conti.cell(row=2, column=2, value="Conto Corrente")
    ws_conti.cell(row=3, column=2, value="Deposito Libero")
    ws_conti.cell(row=4, column=2, value="Deposito Vincolato")

    widths_conti = [25, 20, 15, 18, 18, 20, 30]
    for idx, width in enumerate(widths_conti, 1):
        ws_conti.column_dimensions[get_column_letter(idx)].width = width

    # ========== MUTUO & DEBITI ==========
    ws_mutuo = wb.create_sheet("Mutuo & Debiti")
    mutuo_headers = ["Tipo Debito", "Importo Iniziale (€)", "Debito Residuo (€)",
                    "Tasso Interesse (%)", "Rata Mensile (€)", "Data Inizio",
                    "Scadenza", "Note"]

    for col_idx, header in enumerate(mutuo_headers, 1):
        cell = ws_mutuo.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    widths_mutuo = [20, 20, 20, 18, 18, 15, 15, 30]
    for idx, width in enumerate(widths_mutuo, 1):
        ws_mutuo.column_dimensions[get_column_letter(idx)].width = width

    # ========== STIPENDI ==========
    ws_stipendi = wb.create_sheet("Stipendi")
    stipendi_headers = ["Anno", "Mese", "Stipendio Lordo (€)", "Stipendio Netto (€)",
                       "Bonus/Extra (€)", "Totale Netto (€)", "Note"]

    for col_idx, header in enumerate(stipendi_headers, 1):
        cell = ws_stipendi.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_stipendi.cell(row=2, column=6, value='=IF(D2<>"",D2+E2,"")')

    current_year = datetime.now().year
    months = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
             "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

    for idx, month in enumerate(months, 2):
        ws_stipendi.cell(row=idx, column=1, value=current_year)
        ws_stipendi.cell(row=idx, column=2, value=month)

    widths_stipendi = [10, 15, 20, 20, 18, 18, 30]
    for idx, width in enumerate(widths_stipendi, 1):
        ws_stipendi.column_dimensions[get_column_letter(idx)].width = width

    # ========== SPESE MENSILI ==========
    ws_spese = wb.create_sheet("Spese Mensili")
    spese_headers = ["Anno", "Mese", "Affitto/Mutuo (€)", "Bollette (€)", "Spesa (€)",
                    "Trasporti (€)", "Assicurazioni (€)", "Abbonamenti (€)",
                    "Svago (€)", "Altro (€)", "TOTALE (€)", "Note"]

    for col_idx, header in enumerate(spese_headers, 1):
        cell = ws_spese.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_spese.cell(row=2, column=11, value='=SUM(C2:J2)')

    for idx, month in enumerate(months, 2):
        ws_spese.cell(row=idx, column=1, value=current_year)
        ws_spese.cell(row=idx, column=2, value=month)

    widths_spese = [10, 15, 15, 12, 12, 12, 15, 15, 12, 12, 15, 30]
    for idx, width in enumerate(widths_spese, 1):
        ws_spese.column_dimensions[get_column_letter(idx)].width = width

    # ========== TASSE ==========
    ws_tasse = wb.create_sheet("Tasse")
    tasse_headers = ["Anno Fiscale", "Tipo", "Descrizione", "Base Imponibile (€)",
                    "Aliquota (%)", "Importo Pagato (€)", "Data Pagamento", "Note"]

    for col_idx, header in enumerate(tasse_headers, 1):
        cell = ws_tasse.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_tasse.cell(row=2, column=2, value="Capital Gain")
    ws_tasse.cell(row=3, column=2, value="Dividendi")
    ws_tasse.cell(row=4, column=2, value="Interessi")
    ws_tasse.cell(row=2, column=6, value='=IF(D2<>"",D2*(E2/100),"")')

    widths_tasse = [15, 18, 25, 20, 15, 20, 18, 30]
    for idx, width in enumerate(widths_tasse, 1):
        ws_tasse.column_dimensions[get_column_letter(idx)].width = width

    # ========== STORICO PATRIMONIO ==========
    ws_storico = wb.create_sheet("Storico Patrimonio")
    storico_headers = ["Data", "Investimenti", "Fondi", "Liquidità",
                      "Totale Attivo", "Debiti", "Patrimonio Netto", "Var. % Mensile"]

    for col_idx, header in enumerate(storico_headers, 1):
        cell = ws_storico.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border_thin

    ws_storico.cell(row=2, column=5, value='=SUM(B2:D2)')
    ws_storico.cell(row=2, column=7, value='=E2-F2')
    ws_storico.cell(row=2, column=8, value='=IF(G1<>0,(G2-G1)/G1*100,"")')

    widths_storico = [12, 18, 18, 18, 18, 18, 20, 18]
    for idx, width in enumerate(widths_storico, 1):
        ws_storico.column_dimensions[get_column_letter(idx)].width = width

    # Salva il file
    output_file = "Portfolio_Tracker.xlsx"
    wb.save(output_file)

    print(f"✅ File '{output_file}' creato con successo!")
    print(f"\n📊 NUOVA VERSIONE 2.0 - Design migliorato!")
    print(f"\n🎨 Fogli creati:")
    print(f"  1. Dashboard - Design professionale con KPI e spazio per grafico a torta")
    print(f"  2. Investimenti - NUOVO! Azioni, ETF e Bond unificati")
    print(f"  3. Fondi Comuni - Fondi di investimento")
    print(f"  4. Conti & Liquidità - Cash e depositi")
    print(f"  5. Mutuo & Debiti - Gestione debiti")
    print(f"  6. Stipendi - Storico mensile")
    print(f"  7. Spese Mensili - Tracking spese (opzionale)")
    print(f"  8. Tasse - Capital gain e imposte")
    print(f"  9. Storico Patrimonio - Performance nel tempo")
    print(f"\n🎨 Miglioramenti:")
    print(f"  ✓ Dashboard con layout moderno e KPI in evidenza")
    print(f"  ✓ Palette colori professionale")
    print(f"  ✓ Foglio Investimenti unificato con colonna Tipo")
    print(f"  ✓ Spazio dedicato per grafico a torta nella Dashboard")
    print(f"  ✓ Formattazione migliorata con bordi e colori")
    print(f"\n📥 Importa il file in Google Sheets e aggiungi il grafico a torta!")
    print(f"    (Seleziona celle A9:B11 → Inserisci → Grafico → Torta)")

if __name__ == "__main__":
    create_portfolio_tracker()

# 📊 Portfolio Tracker Excel v2.0

Un file Excel completo e professionale per tracciare il tuo portafoglio finanziario e gestire le finanze personali. **Versione 2.0** con design moderno, palette colori professionale e dashboard migliorata con grafico a torta.

## ✨ Features

- **🎨 Design Professionale v2.0**: Dashboard moderna con KPI in evidenza e palette colori accattivante
- **📊 Grafico a Torta**: Visualizzazione immediata della composizione del portafoglio
- **📈 Tracking Automatico Prezzi**: Integrazione con GOOGLEFINANCE per aggiornamento real-time
- **💼 Foglio Investimenti Unificato**: Azioni, ETF e Bond in un unico foglio con colonna "Tipo"
- **💰 Gestione Completa**: Fondi Comuni, Liquidità, Mutui e Debiti
- **💵 Storico Stipendi**: Tracking mensile di stipendi lordi, netti e bonus
- **📉 Analisi Performance**: Calcolo automatico di P&L, rendimenti percentuali e variazioni
- **🧾 Gestione Tasse**: Tracking capital gain, dividendi e imposte
- **📅 Storico Patrimonio**: Snapshot mensile per analisi trend nel tempo

## 🚀 Quick Start

### 1. Scarica il file

```bash
# Clona il repository
git clone https://github.com/Pablo95byte/PortfolioTracking.git
cd PortfolioTracking

# Oppure scarica direttamente il file
# Portfolio_Tracker.xlsx
```

### 2. Importa in Google Sheets

1. Vai su [Google Sheets](https://sheets.google.com)
2. **File** → **Importa** → **Carica**
3. Seleziona `Portfolio_Tracker.xlsx`
4. Scegli **"Sostituisci foglio di lavoro"** come opzione di importazione

### 3. Inizia a tracciare!

Le formule GOOGLEFINANCE si attiveranno automaticamente. Inizia inserendo i tuoi dati nei vari fogli.

## 📑 Struttura del File

### 1️⃣ Dashboard - NUOVO DESIGN! 🎨

Dashboard professionale con **design moderno v2.0**:

**KPI Principali in evidenza:**
- **PATRIMONIO NETTO** (grande, al centro)
- TOTALE ATTIVO (verde)
- TOTALE PASSIVO (rosso)

**Composizione Portafoglio:**
- Tabella con valori e percentuali per categoria
- **Spazio dedicato per GRAFICO A TORTA** (istruzioni incluse nel file)
- Dettaglio investimenti suddiviso per Azioni, ETF, Bond

**Caratteristiche:**
- Palette colori professionale (blu, verde, arancione)
- Bordi eleganti e celle colorate
- Layout pulito e moderno
- Tutti i valori si aggiornano automaticamente

### 2️⃣ Investimenti - UNIFICATO! 💼

**NUOVO!** Azioni, ETF e Bond in un unico foglio per una gestione più semplice.

**Puoi aggiungere QUANTE RIGHE VUOI per ogni tipo** - il file include già 8 righe di esempio (3 azioni, 3 ETF, 2 bond) ma puoi aggiungerne infinite!

**Colonne:**
- **Tipo** (Azione/ETF/Bond) - Con colori diversi per tipo
- Ticker/ISIN (es. `AAPL`, `BIT:ENI`, `EPA:IWDA`)
- Nome
- Quantità/Nominale
- Prezzo Carico e **Prezzo Corrente** (automatico via GOOGLEFINANCE)
- Valore Carico e Valore Corrente
- **P&L** (€ e %)
- Dividendi/Cedole ricevuti
- Data Acquisto
- Scadenza (per bond)
- Note

**Righe di esempio già incluse nel file:**

**Azioni** (puoi averne quante vuoi):
```
Azione | AAPL      | Apple Inc.
Azione | BIT:ENI   | Eni S.p.A.
Azione | BIT:UCG   | UniCredit
```

**ETF** (puoi averne quanti vuoi):
```
ETF    | VOO       | Vanguard S&P 500 ETF
ETF    | EPA:IWDA  | iShares Core MSCI World
ETF    | AMS:VWCE  | Vanguard FTSE All-World
```

**Bond** (puoi averne quanti vuoi):
```
Bond   | IT0005423745 | BTP Italia 2030
Bond   |              | BTP 2.5% 2033
```

**Vantaggi del foglio unificato:**
- Vista completa di tutti gli investimenti in un colpo d'occhio
- **Aggiungi tutte le righe che vuoi** - ogni asset nella sua riga
- Filtri e ordinamenti più semplici (filtra per "Tipo" per vedere solo ETF, o solo Azioni)
- Confronto diretto tra asset class
- Colori automatici per distinguere i tipi (copia una riga esistente per mantenere i colori)

### 3️⃣ Fondi Comuni

Tracking fondi di investimento.

**Campi:**
- Nome fondo e ISIN
- Numero quote
- Prezzo di carico e corrente
- Valore totale
- P&L (€ e %)

### 4️⃣ Conti & Liquidità

Gestione conti correnti, depositi vincolati e liquidità.

**Tipi di conto:**
- Conto Corrente
- Deposito Libero
- Deposito Vincolato

**Campi:**
- Nome banca/conto
- Saldo attuale
- Tasso di interesse
- Scadenza vincolo (se applicabile)
- Interessi maturati

### 5️⃣ Mutuo & Debiti

Tracking di mutui, prestiti e altri debiti.

**Campi:**
- Tipo debito
- Importo iniziale
- **Debito residuo**
- Tasso di interesse
- Rata mensile
- Date inizio e scadenza

### 6️⃣ Stipendi

Storico mensile degli stipendi (pre-popolato con i 12 mesi dell'anno corrente).

**Campi:**
- Anno e Mese
- Stipendio Lordo
- Stipendio Netto
- Bonus/Extra
- **Totale Netto** (calcolato automaticamente)

### 7️⃣ Spese Mensili (Opzionale)

Tracking delle spese ricorrenti mensili.

**Categorie:**
- Affitto/Mutuo
- Bollette
- Spesa alimentare
- Trasporti
- Assicurazioni
- Abbonamenti
- Svago
- Altro
- **TOTALE** (calcolato automaticamente)

### 8️⃣ Tasse

Gestione delle tasse su investimenti.

**Tipi di tasse:**
- Capital Gain (vendita titoli)
- Dividendi
- Interessi (conti deposito, bond)

**Campi:**
- Anno fiscale
- Tipo e descrizione
- Base imponibile
- Aliquota (%)
- Importo pagato (calcolato automaticamente)
- Data pagamento

### 9️⃣ Storico Patrimonio

Snapshot mensile del patrimonio per analisi performance nel tempo.

**Dati tracciati:**
- Valore Investimenti (totale Azioni+ETF+Bond)
- Fondi Comuni
- Liquidità
- Totale Attivo
- Debiti
- **Patrimonio Netto**
- Variazione % mensile

*Da compilare manualmente a fine mese per creare grafici storici.*

## 🔧 Come Usare GOOGLEFINANCE

### Sintassi Base

```excel
=GOOGLEFINANCE(ticker, [attributo])
```

### Esempi Pratici

**Azioni USA:**
```excel
=GOOGLEFINANCE("AAPL", "price")          → Prezzo Apple
=GOOGLEFINANCE("MSFT", "price")          → Prezzo Microsoft
```

**Azioni Italiane:**
```excel
=GOOGLEFINANCE("BIT:ENI", "price")       → Prezzo ENI
=GOOGLEFINANCE("BIT:UCG", "price")       → Prezzo UniCredit
=GOOGLEFINANCE("BIT:ISP", "price")       → Prezzo Intesa Sanpaolo
```

**ETF Europei:**
```excel
=GOOGLEFINANCE("EPA:IWDA", "price")      → iShares MSCI World (Parigi)
=GOOGLEFINANCE("AMS:VWCE", "price")      → Vanguard All-World (Amsterdam)
=GOOGLEFINANCE("LON:VUSA", "price")      → Vanguard S&P 500 (Londra)
```

### Codici Borsa Comuni

| Codice | Borsa |
|--------|-------|
| `NASDAQ:` o nessuno | NASDAQ (USA) |
| `NYSE:` | New York Stock Exchange |
| `BIT:` | Borsa Italiana (Milano) |
| `EPA:` | Euronext Paris |
| `AMS:` | Euronext Amsterdam |
| `LON:` | London Stock Exchange |
| `FRA:` | Frankfurt Stock Exchange |

## 💡 Consigli d'Uso

### Aggiornamento Prezzi

- I prezzi GOOGLEFINANCE si aggiornano automaticamente durante le ore di mercato
- Per forzare l'aggiornamento: modifica una cella e premi Invio
- I prezzi potrebbero avere un ritardo di ~15 minuti

### Backup e Versioning

- Crea copie mensili del file per tenere uno storico
- Usa Google Drive per backup automatici cloud
- Considera di esportare periodicamente in Excel (.xlsx) per backup locale

### Ottimizzazione Performance

- Evita di inserire troppi ticker nella stessa riga (max ~50-100 GOOGLEFINANCE per foglio)
- Se il file rallenta, riduci il numero di formule GOOGLEFINANCE
- Per bond e fondi comuni, inserisci i prezzi manualmente (spesso non disponibili su GOOGLEFINANCE)

### Privacy e Sicurezza

- **NON condividere pubblicamente** il file compilato (contiene dati finanziari personali)
- Usa le impostazioni di condivisione Google Sheets per limitare l'accesso
- Considera l'uso di autenticazione a due fattori su Google Account

## 📊 Esempi di Utilizzo

### Esempio 1: Aggiungere un'azione

1. Vai al foglio **Investimenti**
2. Inserisci nella prima riga vuota:
   - **Tipo**: `Azione` (o `ETF` o `Bond`)
   - **Ticker**: `AAPL`
   - **Nome**: `Apple Inc.`
   - **Quantità**: `10`
   - **Prezzo Carico**: `150.00`
   - **Data Acquisto**: `01/01/2024`
3. Il prezzo corrente si aggiornerà automaticamente via GOOGLEFINANCE
4. P&L verrà calcolato automaticamente
5. La cella "Tipo" si colorerà automaticamente (blu per Azioni, verde per ETF, arancione per Bond)

### Esempio 2: Tracciare uno stipendio

1. Vai al foglio **Stipendi**
2. Trova il mese corrente (già pre-popolato)
3. Inserisci:
   - **Stipendio Lordo**: `3000`
   - **Stipendio Netto**: `2100`
   - **Bonus/Extra**: `500` (se applicabile)
4. Il totale netto verrà calcolato automaticamente (2600€)

### Esempio 3: Aggiungere un mutuo

1. Vai al foglio **Mutuo & Debiti**
2. Inserisci:
   - **Tipo Debito**: `Mutuo Prima Casa`
   - **Importo Iniziale**: `200000`
   - **Debito Residuo**: `180000`
   - **Tasso Interesse**: `2.5`
   - **Rata Mensile**: `850`
   - **Scadenza**: `01/01/2040`
3. Il valore apparirà automaticamente nel Dashboard come passivo

## 🔄 Aggiornamento Mensile

Routine suggerita per mantenere il tracker aggiornato:

1. **Inizio mese**: Inserisci stipendio del mese precedente
2. **Metà mese**: Verifica i prezzi di mercato e dividendi ricevuti
3. **Fine mese**:
   - Aggiorna il foglio **Storico Patrimonio** con snapshot corrente
   - Inserisci spese mensili (se usi il foglio opzionale)
   - Verifica debito residuo del mutuo
   - Controlla tasse pagate (capital gain, dividendi)

## 🛠️ Personalizzazione

### Aggiungere Nuove Colonne

Puoi personalizzare i fogli aggiungendo colonne extra, ad esempio:
- Settore/Industria (per azioni)
- Rating ESG
- Target price
- Peso % nel portafoglio
- Commissioni pagate

### Aggiungere Grafici

Google Sheets permette di creare grafici facilmente:

1. Seleziona i dati del foglio **Storico Patrimonio**
2. **Inserisci** → **Grafico**
3. Scegli un grafico a linee per visualizzare l'andamento del patrimonio nel tempo

### Creare Pivot Tables

Per analisi avanzate, usa le tabelle pivot:

- Analizza performance per settore
- Confronta rendimenti per anno
- Calcola allocation % per asset class

## 🐛 Risoluzione Problemi

### GOOGLEFINANCE non funziona

**Problema**: La formula restituisce `#N/A` o `#ERROR`

**Soluzioni**:
- Verifica che il ticker sia corretto
- Prova ad aggiungere il codice borsa (es. `BIT:ENI` invece di `ENI`)
- Alcuni titoli potrebbero non essere disponibili su Google Finance
- Inserisci manualmente il prezzo se il ticker non è supportato

### Il file è lento

**Problema**: Google Sheets rallenta con molte formule

**Soluzioni**:
- Riduci il numero di formule GOOGLEFINANCE (max ~50-100 per foglio)
- Copia e incolla come valori i prezzi storici non più necessari
- Dividi il portafoglio in più file (es. un file per anno)

### Prezzi non aggiornati

**Problema**: I prezzi sembrano vecchi

**Soluzioni**:
- GOOGLEFINANCE aggiorna durante le ore di mercato (con ritardo ~15 min)
- Forza l'aggiornamento modificando una cella e premendo Invio
- Chiudi e riapri il file

## 📝 Rigenerare il File

Se vuoi rigenerare il file Excel da zero:

```bash
# Installa le dipendenze
pip install openpyxl

# Esegui lo script
python3 create_portfolio_tracker.py

# Verrà creato Portfolio_Tracker.xlsx
```

## 🤝 Contributi

Contributi, suggerimenti e segnalazioni di bug sono benvenuti!

1. Fai un fork del repository
2. Crea un branch per la tua feature (`git checkout -b feature/nuova-funzionalita`)
3. Committa le modifiche (`git commit -m 'Aggiungi nuova funzionalità'`)
4. Pusha il branch (`git push origin feature/nuova-funzionalita`)
5. Apri una Pull Request

## 📄 Licenza

Questo progetto è distribuito con licenza libera. Sentiti libero di usarlo, modificarlo e condividerlo.

## ⚠️ Disclaimer

Questo strumento è fornito solo a scopo informativo e organizzativo. Non costituisce consulenza finanziaria. Verifica sempre i dati con le tue fonti ufficiali e consulta un professionista per decisioni di investimento.

---

**Creato con ❤️ per gestire le finanze personali in modo semplice ed efficace.**

Per domande o supporto, apri una [Issue](https://github.com/Pablo95byte/PortfolioTracking/issues) su GitHub.

# EV Field Service – Pricing & Margini (Streamlit)

App Streamlit per:
- caricare **listini Partner** (uno per regione, stesso formato del template Excel)
- caricare **prezzario Cliente**
- confrontare prezzi e calcolare **margine** su N installazioni
- calcolare **rebate al cliente finale** (default 5%)
- esportare un report Excel.

## Struttura repo
```
.
├─ app.py
├─ requirements.txt
├─ data/
│  ├─ partners/          # listini partner salvati per regione (file .xlsx)
│  └─ defaults/          # opzionale (puoi aggiungere template o esempi)
└─ .streamlit/
   └─ config.toml
```

## Avvio locale
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy su Streamlit Community Cloud (via GitHub)
1. Crea un repo su GitHub e carica questi file.
2. Vai su Streamlit Community Cloud e collega il repo.
3. Entry point: `app.py`.

> Nota: su Streamlit Cloud il filesystem può essere effimero: i file caricati potrebbero non persistere.  
> Per uso production, salva i listini su storage esterno (S3, Azure Blob, GDrive, DB) e caricali via API.

## Formato file atteso
- Il file Excel deve seguire il template:
  - riga `Installazione ...`
  - intestazioni distanze in riga (colonne C..)
  - righe `Item ...` in colonna B e prezzi nelle colonne distanza.
- Override CSV (opzionale): colonne `block,distance,item_id,fixed_price`

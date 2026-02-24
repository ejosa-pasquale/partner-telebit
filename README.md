# EV Field Service – Pricing & Margini (Streamlit)

## Novità v2
- **Prezzario cliente precaricabile** nel repo: `data/defaults/client_pricelist.xlsx`
- **Listini partner persistenti**: salvati in `data/partners/<Regione>.xlsx` (riutilizzabili senza ricaricare)

> Nota: Streamlit Community Cloud può avere filesystem non persistente. Per produzione: storage esterno (S3/Blob/DB).

## Avvio locale
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Precaricare prezzario cliente
1. Metti il tuo file Excel nel repo:
   - `data/defaults/client_pricelist.xlsx`
2. L'app mostrerà l'opzione **Usa precaricato (repo)**.

## Override prezzi partner (opzionale)
CSV con colonne: `block,distance,item_id,fixed_price`

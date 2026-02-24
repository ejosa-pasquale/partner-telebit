# EV Field Service – Pricing & Margini (Streamlit) – v3

## Novità v3
- **Override totale partner** (CSV)
  - colonne: `block,distance,partner_total_override`
  - opzionale: `region`
- Prezzario cliente precaricabile: `data/defaults/client_pricelist.xlsx`
- Listini partner persistenti: `data/partners/<Regione>.xlsx`

## Avvio locale
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Override totale partner
Se carichi un override totale, il **Totale Partner (unitario)** viene preso dal CSV e **non** dalla somma item.

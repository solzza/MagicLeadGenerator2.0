# Magic Leads Generator Streamlit

Lokal ersattare for Excel/VBA-flodet.

## Starta

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/streamlit run app.py
```

Appen oppnas i webblasaren. Ladda upp Evaboot-filen, valj flik och exportera en Upsales-import.

## Dataflode

- Radata lases fran `.xlsx`, `.xlsm`, `.xls` eller `.csv`.
- Bolagsregistret lases fran `Domaner.csv` i projektmappen.
- Appen matchar bolag exakt, stegvis pa kortare bolagsnamn och till sist med fuzzy match.
- E-post anvands fran importfilen nar den ar verifierad/saker, annars byggs den som `fornamn.efternamn@domain`.
- Osakra rader flaggas i `Issues` och kan granskas innan export.
- Historik sparas lokalt och ska inte commitas.

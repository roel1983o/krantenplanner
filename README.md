# Krantenplanner V1.1

Webtool die:
1) Kordiam Report (xlsx) verwerkt (DEF1)
2) Krantenplanning genereert (DEF2)
3) Hand-out PDF genereert (DEF3)

## Wat upload je per dag?
- **Kordiam Report** (xlsx)
- **Posities en kenmerken** (xlsx)

Alle andere bestanden zitten als vaste **assets** in de repo (zie `assets/`).

## Run lokaal (Docker, aanbevolen)
```bash
docker build -t krantenplanner .
docker run --rm -p 8000:8000 krantenplanner
```
Ga naar: `http://localhost:8000`

## Run lokaal (zonder Docker)
Let op: WeasyPrint heeft OS-libraries nodig (cairo/pango). Op Linux kun je dit via apt installeren.
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```
Ga naar: `http://localhost:8000`

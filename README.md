# Tabbycat Adjudicator Tab Exporter

Web service that reads feedback and pairings from Tabbycat through the REST API (GET only)
and builds an Excel adjudicator tab: average score per round, number of feedback per round,
role colours (chair / panellist / trainee), test score, overall average, final rank and breaking.

## Needs
- Tabbycat URL, tournament slug
- API token of an ADMIN / adjudication-core account (feedback and pairings are not public)

## Colours
Chair `#b10095`, Panellist `#36b6c1`, Trainee `#e96e20`, no feedback grey.

## Deploy to Render
1. Push this folder to its own GitHub repo
2. Create a Web Service on Render
3. Build: `pip install -r requirements.txt`
4. Start: `gunicorn app:app --workers 1 --threads 2 --timeout 120`

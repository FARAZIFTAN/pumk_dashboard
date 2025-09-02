# PUMK Dashboard (Flask)

Small Flask app to display nominative reports for PUMK distribution.

Quick start
- Create a virtual environment and install dependencies:

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1; python -m pip install -r requirements.txt
```

- Create a `.env` from `.env.example` and fill DB creds.
- Run the app:

```powershell
python app.py
```

Notes
- Move secrets to environment variables in production.
- The app uses Bootstrap for a modern UI and supports basic search and pagination.

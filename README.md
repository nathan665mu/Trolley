Trolley.co.uk Scraper – Web App
================================

This repo contains a simple Flask web app that wraps the existing Trolley.co.uk scraper. You can upload an Excel, pick the column with product names, and export matched results to CSV.

Quick start (local)
-------------------

1) Install dependencies

```
python -m pip install -r requirements.txt
```

2) Run the web app

```
python app.py
```

The app will start on http://127.0.0.1:5000

Usage
-----

- Step 1: Upload an .xlsx/.xls file
- Step 2: Choose the column containing product names
- Step 3: Choose how many rows to process, run, and download the CSV

Notes
-----

- The web app caps an "All" run to the first 100 rows to avoid long blocking runs in the browser. Increase with background jobs if needed.
- Please respect Trolley’s terms of service. Consider adding delays for larger runs.

Run with Flask CLI (optional)
----------------------------

```
python -m flask --app app run --host 0.0.0.0 --port 5000
```

Production (Gunicorn, optional)
--------------------------------

```
python -m pip install gunicorn
gunicorn -w 2 -b 0.0.0.0:5000 app:app
```

Docker (optional)
-----------------

The provided Dockerfile is empty. If you want a container, use this minimal image:

```
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 5000
CMD ["python", "app.py"]
```

Then build and run:

```
docker build -t trolley-web .
docker run --rm -p 5000:5000 trolley-web
```


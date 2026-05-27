# ADS — Assignment Document System

A web app that generates personalized stats assignments. Enter your **Name**, **PID**, and **Program** — the system replaces the "Submitted By" details on the front page and generates a pixel-perfect PDF for download.

## How It Works

1. Upload your DOCX template to the `documents/` folder
2. Configure the placeholder text in `documents/settings.json`
3. Users visit the site, enter their details, and get a customized PDF instantly

## Tech Stack

- **Backend:** Python / Flask
- **PDF Conversion:** LibreOffice (headless)
- **DOCX Processing:** python-docx
- **Frontend:** HTML + Three.js (3D sound wave background)
- **Deployment:** Docker → Render

## Run Locally

```bash
docker build -t ads-system .
docker run -p 5000:5000 ads-system
```

Then visit `http://localhost:5000`

## Deploy on Render

1. Push this repo to GitHub
2. Go to [Render Dashboard](https://dashboard.render.com)
3. Click **New → Web Service**
4. Connect this GitHub repo
5. Set **Environment** to **Docker**
6. Deploy — done!

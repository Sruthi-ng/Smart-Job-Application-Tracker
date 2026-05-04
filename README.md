# Smart Job Application Tracker

An AI-powered web application that extracts structured data from job descriptions and logs everything into Google Sheets — built for teams and individuals who want to track job applications effortlessly.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30+-FF4B4B?logo=streamlit&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-Connected-34A853?logo=google-sheets&logoColor=white)
![Gemini](https://img.shields.io/badge/Gemini%20AI-Integrated-4285F4?logo=google&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

---

## What It Does

1. **Privacy First (Bring Your Own Sheet)**: Users provide their own Google Sheet URL when they open the app. Their data is saved securely to their personal sheet, keeping applications 100% private.
2. **AI Extraction**: Users paste a job description (or upload a PDF/DOCX/TXT). AI extracts the company name, job title, job ID, and other details automatically.
3. **Developer Analytics**: While user data remains private, the app logs usage metrics (Timestamp, User Identifier, Event) to a hidden Admin Google Sheet so the developer can track app usage.
4. **Dashboard**: The user's dashboard shows all their applications with KPI cards (total, applied, interviews, offers).

> If AI is unavailable (quota exceeded, offline), the app automatically falls back to local regex extraction — it always works.

---

## Tech Stack

| Layer | Technology | Purpose |
|---|---|---|
| Frontend | Streamlit | Web UI framework |
| AI | Google Gemini (free tier) | Structured data extraction from text |
| Database | Google Sheets | Shared, cloud-based storage |
| Parsing | PyPDF2, python-docx | PDF and DOCX text extraction |
| Auth | Google Service Account | Secure API access |

**Total cost: $0** — all services used are within free tiers.

---

## Quick Start

### Prerequisites

- Python 3.10+
- A Google account
- A Google Cloud project (free)

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR_ORG/smart-job-tracker.git
cd smart-job-tracker
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Set Up Credentials

You need three things:

| Credential | Where to Get It |
|---|---|
| **Gemini API Key** | [aistudio.google.com/apikey](https://aistudio.google.com/apikey) — click "Create API Key" |
| **Google Service Account** | [Google Cloud Console](https://console.cloud.google.com) — see detailed steps below |
| **Google Sheet URL** | Create a new sheet at [sheets.google.com](https://sheets.google.com) |

#### Google Service Account Setup

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project (or use an existing one)
3. Enable **Google Sheets API** and **Google Drive API**
4. Go to **IAM & Admin > Service Accounts** > Create service account
5. Give it a name (e.g., `job-tracker-bot`) and the **Editor** role
6. Go to the service account > **Keys** > **Add Key** > **Create new key** > **JSON**
7. Download the JSON file — you'll use it in the next step

#### Google Sheet Setup

1. Create a new Google Sheet
2. Add these headers in Row 1:

   | A | B | C | D | E | F | G |
   |---|---|---|---|---|---|---|
   | Company | Job Title | Job ID | Date Applied | Status | Comments | Submitted By |

3. **Share the sheet** with your service account's `client_email` (from the JSON file) as **Editor**
4. Copy the full sheet URL

> **Adding more columns**: Simply add new column headers in your Google Sheet. The app reads columns dynamically — they'll appear in the dashboard automatically.

### 4. Configure Secrets

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Edit `.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "your-gemini-api-key"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/your-sheet-id/edit"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-bot@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
```

### 5. Run Locally

```bash
python -m streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Deploy to Streamlit Community Cloud (Free)

This makes the app accessible to anyone with a link — no installation needed.

### 1. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_ORG/smart-job-tracker.git
git push -u origin main
```

> **Security**: `.streamlit/secrets.toml` is in `.gitignore` — credentials are never pushed to GitHub.

### 2. Deploy

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click **New app**
3. Select your GitHub repo, branch `main`, file `app.py`
4. Click **Advanced settings** > paste your `secrets.toml` contents into the Secrets field
5. Click **Deploy**

Your app is live at `https://your-app-name.streamlit.app`.

### 3. Share

Send the URL to anyone. They open it in a browser, paste a job description, and submit. No accounts, no installs.

---

## Customisation

### Adding Columns

Add any column header to Row 1 of your Google Sheet (e.g., `Location`, `Salary Range`, `Source`). The app dashboard will display them automatically on the next page refresh.

### User Tracking

Every submission includes a "Submitted By" field. This logs who added each entry — useful for teams sharing a single tracker.

### Modifying AI Extraction

The AI prompt is defined in `EXTRACTION_PROMPT` inside `app.py`. You can modify it to extract additional fields or change the extraction behaviour.

---

## Project Structure

```
smart-job-tracker/
├── .gitignore                     # Protects secrets from being committed
├── .streamlit/
│   ├── secrets.toml               # Your credentials (never committed)
│   └── secrets.toml.example       # Template for new contributors
├── app.py                         # Main application
├── requirements.txt               # Python dependencies
└── README.md                      # This file
```

---

## Architecture

This app uses a **Multi-Tenant BYO-Sheet** architecture:

1. **Admin Database**: Configured in `.streamlit/secrets.toml`. Used purely for logging usage analytics (a new `Analytics` tab is created automatically).
2. **User Database**: Configured at runtime via the UI. End-users create and connect their own Google Sheet. Their extracted job data goes *only* to their sheet.

```text
User (Browser)
     │
     ├── 1. Connects personal Google Sheet URL
     ├── 2. Pastes Job Description
     │
     ▼
Streamlit App (app.py)
     │
     ├──► Google Gemini API ──► Structured JSON
     │    (fallback: regex)
     │
     ├──► User's Google Sheet ──► Reads/Writes Job Data (100% Private)
     │
     └──► Admin Google Sheet ──► Appends to `Analytics` tab (Usage Tracking)
```

---

## FAQ

**Q: Does this cost anything?**
No. Gemini free tier (1,500 requests/day), Google Sheets API, and Streamlit Community Cloud are all free.

**Q: What if the AI quota runs out?**
The app automatically falls back to local regex extraction. You can edit the fields before saving.

**Q: Can multiple people use this simultaneously?**
Yes. Google Sheets handles concurrent writes. Each user's name is logged in the "Submitted By" column.

**Q: Can I add more columns?**
Yes. Add column headers in the Google Sheet — the dashboard reflects them automatically.

**Q: Is my data secure?**
API keys are stored in Streamlit's secrets manager (never in the code). The Google Sheet is only accessible to people you share it with.

---

## License

MIT License — use freely, modify as needed.

---

*Built by [WillowVibe Digital Solutions](https://github.com/YOUR_ORG)*

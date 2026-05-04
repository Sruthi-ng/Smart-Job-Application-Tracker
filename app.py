"""
Smart Job Application Tracker
==============================
A Streamlit application that extracts structured data from job descriptions
(pasted text or uploaded files) using Google Gemini AI, then logs everything
into a Google Sheet for easy tracking.

Features:
  - Paste text or upload PDF/DOCX/TXT files
  - AI extraction with local regex fallback (always works, zero cost)
  - Dynamic columns — add/remove columns in Google Sheets freely
  - User tracking — logs who submitted each entry
  - One-click access via Streamlit Cloud URL

Author : WillowVibe Digital Solutions
Stack  : Streamlit, Google Sheets (gspread), Google Gemini (google-genai),
         PyPDF2, python-docx
"""

import json
import re
import datetime
import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
from google import genai


# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart Job Tracker",
    page_icon=":briefcase:",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body { font-family: 'Inter', sans-serif; }
/* Make sure we don't override Streamlit's internal icon fonts */
.stMarkdown, .stText, .stButton { font-family: 'Inter', sans-serif; }

.main .block-container { padding-top: 1.5rem; max-width: 1000px; }

.hero {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 16px; padding: 2rem 1.5rem; margin-bottom: 1.5rem;
    text-align: center; box-shadow: 0 8px 32px rgba(102,126,234,.35);
}
.hero h1 { color:#fff!important; font-size:2rem; font-weight:800; margin:0; }
.hero p  { color:rgba(255,255,255,.85)!important; margin-top:.4rem; font-size:1rem; }

.kpi {
    background: linear-gradient(135deg,#f8f9ff,#eef1ff);
    border: 1px solid #e0e5ff; border-radius: 14px;
    padding: 1.2rem .8rem; text-align: center;
    transition: transform .2s, box-shadow .2s;
}
.kpi:hover { transform:translateY(-3px); box-shadow:0 6px 20px rgba(102,126,234,.15); }
.kpi-num   { font-size:1.8rem; font-weight:800; color:#667eea; }
.kpi-label { font-size:.75rem; color:#666; text-transform:uppercase; letter-spacing:1px; margin-top:2px; }

.stButton > button {
    background: linear-gradient(135deg,#667eea,#764ba2);
    color:#fff; border:none; border-radius:10px;
    padding:.55rem 1.4rem; font-weight:600; transition:all .25s;
}
.stButton > button:hover {
    transform:translateY(-2px); box-shadow:0 6px 20px rgba(102,126,234,.4);
}

.stTabs [data-baseweb="tab-list"] { gap: 8px; }
.stTabs [data-baseweb="tab"]      { border-radius:8px; padding:8px 16px; font-weight:600; }

/* Sidebar styling (hidden but just in case) */
[data-testid="collapsedControl"] { display: none !important; }
[data-testid="stSidebar"] { display: none !important; }

.help-box {
    background: rgba(255,255,255,.05); border-radius: 10px;
    padding: 1rem; margin-top: .5rem; line-height: 1.6;
}
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS — Document Parsing
# ═════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(file) -> str:
    try:
        reader = PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        st.error(f"Failed to parse PDF: {e}")
        return ""

def extract_text_from_docx(file) -> str:
    try:
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        st.error(f"Failed to parse DOCX: {e}")
        return ""

def extract_text_from_txt(file) -> str:
    try:
        return file.read().decode("utf-8")
    except Exception as e:
        st.error(f"Failed to read TXT: {e}")
        return ""

def parse_uploaded_file(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):   return extract_text_from_pdf(uploaded_file)
    if name.endswith(".docx"):  return extract_text_from_docx(uploaded_file)
    if name.endswith(".txt"):   return extract_text_from_txt(uploaded_file)
    st.error("Unsupported file type.")
    return ""


# ═════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS — Local Regex Fallback
# ═════════════════════════════════════════════════════════════════════════════

def extract_with_regex(raw_text: str) -> dict:
    """
    Best-effort extraction using regex. Works offline, no API, no cost.
    The user can always edit the results before saving.
    """
    today = datetime.date.today().isoformat()
    lines = raw_text.strip().split("\n")

    # ── Company ──
    company = "N/A"
    for pattern in [
        r"(?:company|employer|firma|unternehmen)\s*[:\-–]\s*(.+)",
        r"(?:at|bei|@)\s+([A-Z][A-Za-z\s&.,]+(?:GmbH|Inc|Ltd|Corp|AG|SE|LLC|Co\.?))",
        r"(?:about|ueber|über)\s+([A-Z][A-Za-z\s&.,]+(?:GmbH|Inc|Ltd|Corp|AG|SE|LLC|Co\.?))",
    ]:
        m = re.search(pattern, raw_text, re.IGNORECASE)
        if m:
            company = m.group(1).strip().rstrip(".,;:")
            break
    if company == "N/A":
        for line in lines[:5]:
            line = line.strip()
            if line and len(line) < 60 and not line.lower().startswith(("job", "position", "role", "stelle")):
                company = line.rstrip(".,;:")
                break

    # ── Job Title ──
    title = "N/A"
    for pattern in [
        r"(?:job\s*title|position|role|stelle|jobtitel)\s*[:\-–]\s*(.+)",
        r"(?:hiring|looking\s+for|suchen|gesucht)\s*[:\-–]?\s*(?:a|an|eine[n]?)?\s*(.+)",
    ]:
        m = re.search(pattern, raw_text, re.IGNORECASE)
        if m:
            title = m.group(1).strip().rstrip(".,;:")
            break
    if title == "N/A":
        keywords = ["engineer","developer","manager","analyst","designer","consultant",
                     "scientist","architect","lead","director","specialist","coordinator",
                     "ingenieur","entwickler","berater"]
        for line in lines[:15]:
            lc = line.strip()
            if any(kw in lc.lower() for kw in keywords) and len(lc) < 80:
                title = lc.rstrip(".,;:")
                break

    # ── Job ID ──
    job_id = "N/A"
    for pattern in [
        r"(?:job\s*id|requisition\s*id|req\s*id|reference|ref|stellen-?id|kennziffer)\s*[:\-–#]?\s*([A-Za-z0-9\-_]+)",
        r"(?:id|ref)\s*[:\-–#]\s*([A-Za-z0-9\-_]{3,})",
    ]:
        m = re.search(pattern, raw_text, re.IGNORECASE)
        if m:
            job_id = m.group(1).strip()
            break

    # ── Date ──
    date_applied = today
    for pattern in [r"(\d{4}[-/]\d{2}[-/]\d{2})", r"(\d{2}[-/.]\d{2}[-/.]\d{4})"]:
        m = re.search(pattern, raw_text)
        if m:
            for fmt in ("%Y-%m-%d","%Y/%m/%d","%d.%m.%Y","%d/%m/%Y","%m/%d/%Y"):
                try:
                    date_applied = datetime.datetime.strptime(m.group(1), fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
            break

    # ── Comment (first meaningful line) ──
    comments = ""
    for line in lines[:10]:
        if len(line.strip()) > 30:
            comments = line.strip()[:120]
            break

    return {
        "Company": company,
        "Job Title": title,
        "Job ID": job_id,
        "Date Applied": date_applied,
        "Status": "Applied",
        "Comments": comments,
    }


# ═════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS — AI Extraction (Google Gemini)
# ═════════════════════════════════════════════════════════════════════════════

EXTRACTION_PROMPT = """You are a precise data-extraction assistant.
Analyse the following text (a job description, application receipt, or resume)
and extract ONLY these fields.

Return ONLY a valid JSON object — no markdown fences, no commentary:

{{
  "Company": "<company name or N/A>",
  "Job Title": "<job title or N/A>",
  "Job ID": "<job/requisition ID or N/A>",
  "Date Applied": "<YYYY-MM-DD or {today}>",
  "Status": "Applied",
  "Comments": "<one-line summary>"
}}

--- TEXT ---
{text}
"""


def extract_with_ai(raw_text: str) -> dict:
    """
    Try Gemini first. If it fails for ANY reason, fall back to regex.
    Always returns a dict — never returns None.
    """
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    if not api_key or "your-gemini" in api_key:
        st.info("No Gemini API key — using local extraction. You can edit the fields below.")
        return extract_with_regex(raw_text)

    try:
        client = genai.Client(api_key=api_key)
        prompt = EXTRACTION_PROMPT.format(
            text=raw_text[:12_000],
            today=datetime.date.today().isoformat(),
        )
        response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
        text = response.text.strip()

        if text.startswith("```"):
            text = text.split("\n", 1)[-1]
        if text.endswith("```"):
            text = text[:-3].strip()

        data = json.loads(text)
        for key in ["Company", "Job Title", "Job ID", "Date Applied", "Status", "Comments"]:
            data.setdefault(key, "N/A")
        return data

    except Exception as e:
        err = str(e)
        if "429" in err or "RESOURCE_EXHAUSTED" in err:
            st.warning("Gemini daily quota reached — using local extraction. Quota resets tomorrow.")
        else:
            st.warning(f"AI unavailable — using local extraction instead.")
        return extract_with_regex(raw_text)


# ═════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS — Google Sheets (Dynamic Columns)
# ═════════════════════════════════════════════════════════════════════════════

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Core columns that AI extracts. Additional columns in the sheet are displayed
# in the dashboard but are not required.
CORE_COLUMNS = ["Company", "Job Title", "Job ID", "Date Applied", "Status", "Comments", "Submitted By"]


@st.cache_resource(show_spinner=False)
def get_gsheet_client():
    """Authenticate with Google Sheets using service-account credentials."""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except KeyError:
        return None
    except Exception as e:
        st.error(f"Google Sheets auth failed: {e}")
        return None


def get_admin_sheet_url() -> str:
    return st.secrets.get("SPREADSHEET_URL", "")


def log_analytics(client, admin_url: str, event: str):
    """Log anonymous usage to the developer's admin sheet."""
    if not client or not admin_url:
        return
    try:
        spreadsheet = client.open_by_url(admin_url)
        try:
            ws = spreadsheet.worksheet("Analytics")
        except gspread.exceptions.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title="Analytics", rows=1000, cols=3)
            ws.append_row(["Timestamp", "Event"])
        
        ws.append_row([datetime.datetime.now().isoformat(), event])
    except Exception:
        pass  # Fail silently for analytics


def load_sheet(client, url: str) -> tuple[pd.DataFrame, list[str]]:
    """
    Load all data from the Google Sheet.
    Returns (DataFrame, list_of_headers).
    The headers are read dynamically — whatever columns exist in Row 1.
    """
    try:
        ws = client.open_by_url(url).worksheet("Sheet1")
        all_values = ws.get_all_values()
        if not all_values:
            return pd.DataFrame(columns=CORE_COLUMNS), CORE_COLUMNS
        headers = all_values[0]
        if len(all_values) > 1:
            df = pd.DataFrame(all_values[1:], columns=headers)
        else:
            df = pd.DataFrame(columns=headers)
        return df, headers
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Spreadsheet not found. Check the URL and sharing permissions.")
        return pd.DataFrame(), []
    except Exception as e:
        st.error(f"Failed to load sheet: {e}")
        return pd.DataFrame(), []


def append_to_sheet(client, url: str, data: dict, headers: list[str]) -> bool:
    """
    Append one row to the Google Sheet.
    Uses the sheet's actual headers to build the row — supports dynamic columns.
    """
    try:
        ws = client.open_by_url(url).worksheet("Sheet1")
        if not ws.get_all_values():
            ws.append_row(headers)
        row = [data.get(h, "") for h in headers]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Failed to save to Google Sheet: {e}")
        return False


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN UI
# ═════════════════════════════════════════════════════════════════════════════

# ── Hero Banner ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>Smart Job Application Tracker</h1>
    <p>Paste a job description or upload a document &mdash; AI extracts the details automatically.</p>
</div>
""", unsafe_allow_html=True)

with st.expander("ℹ️ How to Use (Click to expand)"):
    st.markdown("""
    **Step 1** — Copy a job description from any website  
    **Step 2** — Paste it in the "Paste Job Description" tab below  
    **Step 3** — Click "Analyse" and review the extracted details  
    **Step 4** — Enter your name and click "Save"  
    
    *If AI is unavailable, the app will automatically fall back to local extraction.*
    """)

# ── Connection Check ─────────────────────────────────────────────────────────
gs_client = get_gsheet_client()
admin_url = get_admin_sheet_url()

config_ok = True
if not gs_client:
    st.error("""
    **⚠️ App Configuration Error!**  
    The developer has not configured the Google Service Account correctly.
    """)
    config_ok = False
    st.stop()

# ── User Setup (BYO Sheet) ───────────────────────────────────────────────────
if "user_sheet_url" not in st.session_state:
    st.markdown("### 🛠️ Welcome! Setup Your Tracker")
    st.info("To keep your data 100% private, you need to connect your own Google Sheet.")
    
    st.markdown("""
    **Follow these 3 easy steps:**
    1. Create a new, blank [Google Sheet](https://sheets.new)
    2. Click **Share** (top right) and share it as an **Editor** with this exact email:
    """)
    
    bot_email = "job-tracker-bot@..."
    try:
        bot_email = st.secrets["gcp_service_account"]["client_email"]
    except:
        pass
        
    st.code(bot_email, language=None)
    st.markdown("3. Copy the URL of your Google Sheet and paste it below:")
    
    user_url = st.text_input("Your Google Sheet URL", placeholder="https://docs.google.com/spreadsheets/d/...")
    
    if st.button("Connect & Start Tracking", type="primary"):
        if not user_url.startswith("http"):
            st.warning("Please enter a valid URL.")
        else:
            with st.spinner("Connecting..."):
                try:
                    # Test connection
                    ws = gs_client.open_by_url(user_url).worksheet("Sheet1")
                    st.session_state["user_sheet_url"] = user_url
                    log_analytics(gs_client, admin_url, "User Connected")
                    st.rerun()
                except Exception as e:
                    st.error("Could not connect. Did you share it with the email address above as an Editor?")
    st.stop()

sheet_url = st.session_state["user_sheet_url"]

# ── Load sheet data (dynamic columns) ────────────────────────────────────────
df = pd.DataFrame()
sheet_headers = CORE_COLUMNS

if config_ok:
    df, sheet_headers = load_sheet(gs_client, sheet_url)
    if not sheet_headers:
        sheet_headers = CORE_COLUMNS

# ── Dashboard: KPI cards + data table ────────────────────────────────────────
if config_ok and not df.empty:
    total     = len(df)
    applied   = int((df["Status"] == "Applied").sum())   if "Status" in df.columns else 0
    interview = int(df["Status"].isin(["Interview","Interviewing"]).sum()) if "Status" in df.columns else 0
    offers    = int((df["Status"] == "Offer").sum())      if "Status" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    for col, num, label in [
        (c1, total, "Total"),
        (c2, applied, "Applied"),
        (c3, interview, "Interviews"),
        (c4, offers, "Offers"),
    ]:
        col.markdown(
            f'<div class="kpi"><div class="kpi-num">{num}</div>'
            f'<div class="kpi-label">{label}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("#### Application Log")
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data as CSV",
        data=csv,
        file_name="job_applications.csv",
        mime="text/csv",
    )

elif config_ok:
    st.info("No applications logged yet. Add one below to get started!")

# ── Input Section ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("#### Add a New Application")

tab_paste, tab_upload, tab_manual = st.tabs([
    "Paste Job Description",
    "Upload a File",
    "Enter Manually",
])

# ── Tab 1: Paste text ────────────────────────────────────────────────────────
with tab_paste:
    st.caption("Copy a job posting from any website and paste it here.")
    pasted = st.text_area(
        "Job description text",
        height=220,
        placeholder="Paste the full job description here...",
        key="paste_area",
        label_visibility="collapsed",
    )
    if st.button("Analyse pasted text", type="primary", use_container_width=True, key="btn_paste"):
        if not pasted or len(pasted.strip()) < 20:
            st.warning("Please paste some text first (at least a few lines).")
        else:
            with st.spinner("Analysing..."):
                result = extract_with_ai(pasted.strip())
            st.session_state["extracted"] = result
            st.session_state["source"] = "paste"
            st.rerun()

# ── Tab 2: Upload file ───────────────────────────────────────────────────────
with tab_upload:
    st.caption("Upload a PDF, DOCX, or TXT file containing a job description.")
    uploaded = st.file_uploader(
        "Choose a file",
        type=["pdf", "docx", "txt"],
        key="file_upload",
        label_visibility="collapsed",
    )
    if uploaded:
        raw = parse_uploaded_file(uploaded)
        if raw and raw.strip():
            if st.button("Analyse uploaded file", type="primary", use_container_width=True, key="btn_upload"):
                with st.spinner("Analysing..."):
                    result = extract_with_ai(raw.strip())
                st.session_state["extracted"] = result
                st.session_state["source"] = "upload"
                st.rerun()
        else:
            st.error("The file appears to be empty or unreadable.")

# ── Tab 3: Manual entry ──────────────────────────────────────────────────────
with tab_manual:
    st.caption("No document? Enter the details by hand.")
    m1, m2 = st.columns(2)
    with m1:
        m_company = st.text_input("Company Name", key="m_company")
        m_title   = st.text_input("Job Title", key="m_title")
        m_jobid   = st.text_input("Job ID (optional)", key="m_jobid")
    with m2:
        m_date    = st.date_input("Date Applied", value=datetime.date.today(), key="m_date")
        m_status  = st.selectbox("Status", ["Applied","Interview","Offer","Rejected","Withdrawn"], key="m_status")
        m_comment = st.text_area("Comments (optional)", key="m_comment", height=68)
    m_user = st.text_input("Your Name", key="m_user", placeholder="Who is submitting this?")

    if st.button("Save manual entry", type="primary", use_container_width=True, key="btn_manual"):
        if not m_company or not m_title:
            st.warning("Please enter at least the Company and Job Title.")
        elif not m_user:
            st.warning("Please enter your name so we know who submitted this.")
        elif not config_ok:
            st.error("Google Sheets not connected.")
        else:
            manual = {
                "Company": m_company, "Job Title": m_title,
                "Job ID": m_jobid or "N/A", "Date Applied": m_date.isoformat(),
                "Status": m_status, "Comments": m_comment,
                "Submitted By": m_user,
            }
            with st.spinner("Saving..."):
                if append_to_sheet(gs_client, sheet_url, manual, sheet_headers):
                    log_analytics(gs_client, admin_url, "Job Logged")
                    st.success("Saved to Google Sheets!")
                    st.balloons()
                    st.cache_resource.clear()
                    st.rerun()


# ── Review & Save (shown after AI/regex extraction) ──────────────────────────
if "extracted" in st.session_state:
    st.markdown("---")
    st.markdown("#### Review Extracted Details")
    st.caption("Edit any field before saving. All fields are editable.")

    ex = st.session_state["extracted"]
    r1, r2 = st.columns(2)
    with r1:
        ex["Company"]    = st.text_input("Company",    value=ex.get("Company",""),   key="rev_company")
        ex["Job Title"]  = st.text_input("Job Title",  value=ex.get("Job Title",""), key="rev_title")
        ex["Job ID"]     = st.text_input("Job ID",     value=ex.get("Job ID",""),    key="rev_jobid")
    with r2:
        ex["Date Applied"] = st.text_input("Date Applied", value=ex.get("Date Applied", datetime.date.today().isoformat()), key="rev_date")
        status_opts = ["Applied","Interview","Offer","Rejected","Withdrawn"]
        cur = ex.get("Status","Applied")
        idx = status_opts.index(cur) if cur in status_opts else 0
        ex["Status"]   = st.selectbox("Status", status_opts, index=idx, key="rev_status")
        ex["Comments"] = st.text_area("Comments", value=ex.get("Comments",""), key="rev_comments", height=68)

    ex_user = st.text_input("Your Name", key="rev_user", placeholder="Who is submitting this?")
    ex["Submitted By"] = ex_user

    col_save, col_cancel = st.columns([3, 1])
    with col_save:
        if st.button("Save to Google Sheets", type="primary", use_container_width=True, key="btn_save"):
            if not ex_user:
                st.warning("Please enter your name.")
            elif not config_ok:
                st.error("Google Sheets not connected.")
            else:
                with st.spinner("Saving..."):
                    if append_to_sheet(gs_client, sheet_url, ex, sheet_headers):
                        log_analytics(gs_client, admin_url, "Job Logged")
                        st.success("Saved to Google Sheets!")
                        st.balloons()
                        del st.session_state["extracted"]
                        st.cache_resource.clear()
                        st.rerun()
    with col_cancel:
        if st.button("Cancel", use_container_width=True, key="btn_cancel"):
            del st.session_state["extracted"]
            st.rerun()


# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<div style="text-align:center;opacity:.4;font-size:.8rem;padding:.5rem 0">'
    'Smart Job Application Tracker &middot; Built with Streamlit &amp; Google Gemini'
    '<br>&copy; 2026 WillowVibe Digital Solutions</div>',
    unsafe_allow_html=True,
)

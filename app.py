# -*- coding: utf-8 -*-
"""
CallAI Analytics - Streamlit App (SaaS Edition)
================================================
Flow:
  1. Silent login to CRM (fixed credentials, no login screen)
  2. Pick client (Weebo / Hari Om Pvt Ltd / F1 INFO SOLUTION), date range
  3. Fetch CDR report (auto-filtered by the selected client's company_id)
  4. Pick call-type: Large (>5 min) / Medium (2-5 min) / Small (<2 min) / Custom
  5. Pick count: All matching, or a manual number
  6. Run VAD (Silero) on the recordings to get real Talk Time / Silence /
     Dead Air / Longest Silence
  7. Download final Excel report with the exact required headers

NOTE: CRM credentials are fixed in this file (no login screen shown to
end users). Nothing here is stored outside this session.
"""

import os
import re
import io
import time
import subprocess
import tempfile
from datetime import date, timedelta
from urllib.parse import urljoin

import numpy as np
import pandas as pd
import requests
import soundfile as sf
from bs4 import BeautifulSoup
import streamlit as st
import torch

# ============================================================
# PAGE CONFIG + SAAS-STYLE THEME
# ============================================================
st.set_page_config(
    page_title="CallAI · Talk-Time Analytics",
    page_icon="📞",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stSidebar"] {display: none;}

    html, body, [class*="css"] {
        font-family: -apple-system, "Segoe UI", Inter, Roboto, Arial, sans-serif;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1100px;
    }

    /* ---- Top hero banner ---- */
    .callai-hero {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
        padding: 28px 32px;
        border-radius: 18px;
        color: white;
        margin-bottom: 28px;
        box-shadow: 0 8px 24px rgba(79, 70, 229, 0.25);
    }
    .callai-hero h1 {
        font-size: 28px;
        font-weight: 700;
        margin: 0 0 4px 0;
        color: white;
    }
    .callai-hero p {
        font-size: 15px;
        margin: 0;
        opacity: 0.9;
    }

    /* ---- Section / step card ---- */
    .step-card {
        background: #FFFFFF;
        border: 1px solid #ECECF4;
        border-radius: 16px;
        padding: 22px 26px;
        margin-bottom: 20px;
        box-shadow: 0 2px 10px rgba(20, 20, 43, 0.04);
    }
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: #4F46E5;
        color: white;
        font-weight: 700;
        font-size: 14px;
        margin-right: 10px;
    }
    .step-title {
        font-size: 17px;
        font-weight: 700;
        color: #14142B;
        display: inline-flex;
        align-items: center;
        margin-bottom: 6px;
    }
    .step-subtitle {
        color: #6E7191;
        font-size: 13.5px;
        margin: 0 0 16px 40px;
    }

    /* ---- Metric pills ---- */
    .metric-pill {
        background: #F5F4FF;
        border: 1px solid #E4E1FF;
        border-radius: 14px;
        padding: 14px 18px;
        text-align: center;
    }
    .metric-pill .value {
        font-size: 24px;
        font-weight: 800;
        color: #4F46E5;
    }
    .metric-pill .label {
        font-size: 12.5px;
        color: #6E7191;
        margin-top: 2px;
    }

    /* ---- Buttons ---- */
    div.stButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
        border: none;
    }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
    }
    div.stDownloadButton > button {
        border-radius: 10px;
        font-weight: 700;
        background: linear-gradient(135deg, #16A34A 0%, #22C55E 100%);
        color: white;
        border: none;
        padding: 0.7rem 1.4rem;
    }

    /* ---- Radio as pill-like segmented control ---- */
    div[role="radiogroup"] label {
        border: 1px solid #E4E1FF;
        padding: 6px 14px;
        border-radius: 20px;
        margin-right: 6px;
    }

    .status-banner-ok {
        background: #ECFDF5;
        border: 1px solid #6EE7B7;
        color: #065F46;
        padding: 10px 16px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 13.5px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callai-hero">
    <h1>📞 CallAI · Talk-Time Analytics</h1>
    <p>Pick a client, fetch calls, filter by duration, and get a real Talk-Time / Silence / Dead-Air report — no technical steps needed.</p>
</div>
""", unsafe_allow_html=True)

CRM_BASE = "https://crmapi.dialdesk.in"
LOGIN_URL = f"{CRM_BASE}/auth/login"
CDR_URL = f"{CRM_BASE}/report/cdr_report"

# ============================================================
# ⚠️ FIXED CRM CREDENTIALS - edit these once, no login screen shown to users
# ============================================================
CRM_EMAIL = "ispark@dialdesk.in"
CRM_PASSWORD = "1234"

# ============================================================
# ⚠️ CLIENTS - name -> company_id (edit this dict to add/remove clients)
# ============================================================
CLIENTS = {
    "Weebo": "687",
    "Hari Om Pvt Ltd": "689",
    "F1 INFO SOLUTION": "609",
}

# ============================================================
# SESSION STATE DEFAULTS
# ============================================================
defaults = {
    "token": None,
    "cdr_df": None,
    "cdr_client": None,
    "final_df": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# CRM FUNCTIONS
# ============================================================
def do_login():
    resp = requests.post(
        LOGIN_URL,
        json={"email": CRM_EMAIL, "password": CRM_PASSWORD},
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        timeout=30,
        proxies=None,
    )
    resp.raise_for_status()
    data = resp.json()
    token = (
        data.get("token")
        or data.get("access_token")
        or (data.get("data", {}) or {}).get("token")
    )
    if not token:
        raise RuntimeError(f"Login response had no token field: {data}")
    st.session_state["token"] = token
    return token

def get_valid_token():
    if not st.session_state.get("token"):
        with st.spinner("Signing in..."):
            do_login()
    return st.session_state["token"]

def fetch_cdr(payload, retry_on_401=True):
    token = get_valid_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    resp = requests.post(CDR_URL, json=payload, headers=headers, timeout=120, proxies=None)
    if resp.status_code == 401 and retry_on_401:
        do_login()
        return fetch_cdr(payload, retry_on_401=False)
    return resp

# ============================================================
# RECORDING DOWNLOAD FUNCTIONS
# ============================================================
def html_recording_to_direct_url(webform_url, retries=3):
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    audio_exts = (".mp3", ".wav", ".m4a", ".mp4")
    for attempt in range(retries):
        try:
            resp = session.get(webform_url, headers=headers, timeout=30)
            resp.raise_for_status()
            if resp.url.lower().endswith(audio_exts):
                return resp.url
            soup = BeautifulSoup(resp.text, "html.parser")
            for tag in soup.find_all(["audio", "video"]):
                src = tag.get("src")
                if src and any(ext in src.lower() for ext in audio_exts):
                    return urljoin(webform_url, src)
            for tag in soup.find_all("source"):
                src = tag.get("src")
                if src and any(ext in src.lower() for ext in audio_exts):
                    return urljoin(webform_url, src)
            for div in soup.find_all(attrs={"data-recording": True}):
                for attr in ["data-recording", "data-url", "data-src", "data-file"]:
                    url = div.get(attr)
                    if url:
                        return urljoin(webform_url, url)
            patterns = [
                r'https?://[^\s"\']+\.(?:mp3|wav|m4a)',
                r'//[^\s"\']+\.(?:mp3|wav|m4a)',
                r'/[^\s"\']+\.(?:mp3|wav|m4a)',
            ]
            for pattern in patterns:
                m = re.search(pattern, resp.text, re.IGNORECASE)
                if m:
                    match = m.group()
                    if match.startswith("//"):
                        return "https:" + match
                    if match.startswith("/"):
                        return urljoin(webform_url, match)
                    return match
            js_patterns = [
                r'recordingUrl\s*[:=]\s*["\']([^"\']+)["\']',
                r'audioUrl\s*[:=]\s*["\']([^"\']+)["\']',
                r'fileUrl\s*[:=]\s*["\']([^"\']+)["\']',
                r'src\s*[:=]\s*["\']([^"\']+\.(?:mp3|wav|m4a))["\']',
            ]
            for pattern in js_patterns:
                m = re.search(pattern, resp.text, re.IGNORECASE)
                if m:
                    return urljoin(webform_url, m.group(1))
            iframe = soup.find("iframe")
            if iframe and iframe.get("src"):
                iframe_src = urljoin(webform_url, iframe.get("src"))
                return html_recording_to_direct_url(iframe_src, retries=retries - 1)
            meta_refresh = soup.find("meta", attrs={"http-equiv": re.compile("refresh", re.I)})
            if meta_refresh and meta_refresh.get("content"):
                m = re.search(r"url=([^;]+)", meta_refresh.get("content"), re.IGNORECASE)
                if m:
                    return html_recording_to_direct_url(urljoin(webform_url, m.group(1)), retries=retries - 1)
            for link in soup.find_all("a", href=True):
                href = link.get("href", "")
                if any(ext in href.lower() for ext in audio_exts):
                    return urljoin(webform_url, href)
            return None
        except Exception:
            if attempt == retries - 1:
                return None
            time.sleep(1)
    return None

def resolve_audio_url(recording_url):
    if not isinstance(recording_url, str) or not recording_url.strip():
        return None
    recording_url = recording_url.strip()
    if recording_url.lower().endswith((".mp3", ".wav", ".m4a", ".mp4")):
        return recording_url
    return html_recording_to_direct_url(recording_url)

# ============================================================
# FLEXIBLE COLUMN MAPPING
# ============================================================
COLUMN_CANDIDATES = {
    "date": ["call_date", "CallDate", "Date"],
    "time": ["start_time", "Time", "StartTime"],
    "agent_name": ["full_name", "AgentName", "agent", "Agent Name"],
    "call_from": ["phone_number", "PhoneNumber", "Call From"],
    "recording": ["Recording", "RecordingUrl", "RecordingURL", "recording_url"],
}

def find_column(df, keys):
    lower_map = {c.lower(): c for c in df.columns}
    for key in keys:
        if key.lower() in lower_map:
            return lower_map[key.lower()]
    return None

def parse_duration_series_to_seconds(series):
    s = series.astype(str).str.strip()
    numeric = pd.to_numeric(s, errors="coerce")
    needs_time_parse = numeric.isna() & s.str.contains(":", na=False)
    if needs_time_parse.any():
        def to_seconds(val):
            parts = val.split(":")
            try:
                parts = [float(p) for p in parts]
            except ValueError:
                return np.nan
            if len(parts) == 3:
                h, m, sec = parts
                return h * 3600 + m * 60 + sec
            elif len(parts) == 2:
                m, sec = parts
                return m * 60 + sec
            return np.nan
        numeric.loc[needs_time_parse] = s.loc[needs_time_parse].apply(to_seconds)
    return numeric

def resolve_duration_column(df):
    """
    Tries every known duration-column candidate and picks the FIRST one that
    actually contains meaningful (non-zero, mostly non-null) data - instead
    of blindly trusting column-name priority, since the same logical field
    can be empty/zero in one column and populated in another depending on
    how the CRM exports it.
    """
    candidates_in_order = [
        ("call_duration", "sec"),
        ("call_duration1", "sec"),
        ("CallDurationSecond", "sec"),
        ("Talkduration", "sec"),
        ("CallDurationMinute", "min"),
    ]
    best_col, best_seconds, best_score = None, None, -1
    for name, unit in candidates_in_order:
        col = find_column(df, [name])
        if not col:
            continue
        seconds = parse_duration_series_to_seconds(df[col])
        if unit == "min":
            seconds = seconds * 60
        non_null = seconds.notna().sum()
        non_zero = (seconds.fillna(0) > 0).sum()
        score = non_zero
        if non_null > 0 and score > best_score:
            best_col, best_seconds, best_score = col, seconds.fillna(0), score
    return best_col, best_seconds

def fmt_hms(total_seconds):
    """Turns raw seconds into a friendly M:SS string for display."""
    if total_seconds is None or (isinstance(total_seconds, float) and np.isnan(total_seconds)):
        return "-"
    total_seconds = int(round(total_seconds))
    m, s = divmod(total_seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"

# ============================================================
# STEP 1 — CLIENT + DATE RANGE + FETCH
# ============================================================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.markdown('<div class="step-title"><span class="step-badge">1</span>Choose Client & Date Range</div>', unsafe_allow_html=True)
st.markdown('<p class="step-subtitle">Only calls belonging to the selected client will be fetched.</p>', unsafe_allow_html=True)

c1, c2 = st.columns([1.2, 1.8])
with c1:
    client_name = st.selectbox("Client", options=list(CLIENTS.keys()))
    company_id = CLIENTS[client_name]
with c2:
    today = date.today()
    date_range = st.date_input(
        "Date range", value=(today - timedelta(days=1), today), max_value=today
    )

if isinstance(date_range, tuple) and len(date_range) == 2:
    from_date, to_date = date_range
else:
    from_date = to_date = date_range

fetch_clicked = st.button("📥  Fetch Calls", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# FETCH CDR REPORT
# ============================================================
if fetch_clicked:
    try:
        payload = {
            "from_date": from_date.strftime("%Y-%m-%d"),
            "to_date": to_date.strftime("%Y-%m-%d"),
            "company_id": str(company_id),
        }
        with st.spinner(f"Fetching calls for {client_name}..."):
            resp = fetch_cdr(payload)
        if resp.status_code == 200:
            data = resp.json()
            records = data.get("data", data) if isinstance(data, dict) else data
            if isinstance(records, dict):
                for v in records.values():
                    if isinstance(v, list):
                        records = v
                        break
            cdr_df = pd.DataFrame(records)
            st.session_state["cdr_df"] = cdr_df
            st.session_state["cdr_client"] = client_name
            st.session_state["final_df"] = None
            if len(cdr_df) == 0:
                st.warning(f"No calls found for **{client_name}** in this date range.")
            else:
                st.markdown(
                    f'<span class="status-banner-ok">✅ Fetched {len(cdr_df)} calls for {client_name}</span>',
                    unsafe_allow_html=True,
                )
        else:
            st.error(f"Fetch failed: HTTP {resp.status_code} — {resp.text[:300]}")
    except Exception as e:
        st.error(f"Fetch error: {e}")

# ============================================================
# STEP 2 — FILTER BY CALL DURATION
# ============================================================
have_data = (
    st.session_state["cdr_df"] is not None
    and len(st.session_state["cdr_df"]) > 0
    and st.session_state.get("cdr_client") == client_name
)

if have_data:
    cdr_df = st.session_state["cdr_df"].copy()
    col_date = find_column(cdr_df, COLUMN_CANDIDATES["date"])
    col_time = find_column(cdr_df, COLUMN_CANDIDATES["time"])
    col_agent = find_column(cdr_df, COLUMN_CANDIDATES["agent_name"])
    col_phone = find_column(cdr_df, COLUMN_CANDIDATES["call_from"])
    col_recording = find_column(cdr_df, COLUMN_CANDIDATES["recording"])
    dur_source_col, duration_seconds = resolve_duration_column(cdr_df)

    if dur_source_col is None:
        st.error(
            "Could not find a usable call-duration column in the CDR response. "
            "Columns present: " + ", ".join(cdr_df.columns)
        )
        st.stop()

    cdr_df["_duration_sec"] = duration_seconds

    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-badge">2</span>Pick Call Type & Count</div>', unsafe_allow_html=True)
    st.markdown('<p class="step-subtitle">Choose which calls you want in the report.</p>', unsafe_allow_html=True)

    # ---- summary pills ----
    p1, p2, p3 = st.columns(3)
    with p1:
        st.markdown(f'<div class="metric-pill"><div class="value">{len(cdr_df)}</div><div class="label">Total calls fetched</div></div>', unsafe_allow_html=True)
    with p2:
        avg_dur = cdr_df["_duration_sec"].mean() if len(cdr_df) else 0
        st.markdown(f'<div class="metric-pill"><div class="value">{fmt_hms(avg_dur)}</div><div class="label">Average call duration</div></div>', unsafe_allow_html=True)
    with p3:
        st.markdown(f'<div class="metric-pill"><div class="value">{client_name}</div><div class="label">Client</div></div>', unsafe_allow_html=True)

    st.write("")
    bcol, ccol = st.columns([2, 1.2])
    with bcol:
        bucket = st.radio(
            "Call type",
            ["All calls", "Large (> 5 min)", "Medium (2 – 5 min)", "Small (< 2 min)", "Custom"],
            horizontal=True,
        )
    with ccol:
        count_mode = st.radio("How many calls?", ["All matching", "Manual number"], horizontal=True)

    # Custom thresholds only appear if "Custom" chosen — everything else uses sensible fixed defaults
    small_max, large_min = 120, 300  # 2 min, 5 min
    if bucket == "Custom":
        t1, t2 = st.columns(2)
        with t1:
            small_max = st.number_input("Small call: below (sec)", min_value=1, value=120, step=10)
        with t2:
            large_min = st.number_input("Large call: above (sec)", min_value=1, value=300, step=10)
        if small_max >= large_min:
            st.error("⚠️ 'Small below' must be smaller than 'Large above'.")

    manual_n = None
    if count_mode == "Manual number":
        manual_n = st.number_input("Number of calls", min_value=1, value=50, step=1)

    # Bucket logic — Medium always fills the gap between small_max and large_min,
    # so no call is ever left uncategorised. "Custom" just redefines the thresholds
    # and shows everything split across Small / Medium / Large using those numbers.
    if bucket.startswith("Large"):
        matched = cdr_df[cdr_df["_duration_sec"] > large_min]
    elif bucket.startswith("Medium"):
        matched = cdr_df[(cdr_df["_duration_sec"] >= small_max) & (cdr_df["_duration_sec"] <= large_min)]
    elif bucket.startswith("Small"):
        matched = cdr_df[cdr_df["_duration_sec"] < small_max]
    else:
        matched = cdr_df  # "All calls" or "Custom" (custom just changes the thresholds above)

    available = len(matched)
    if count_mode == "Manual number":
        selected_df = matched.head(int(manual_n))
        if available < manual_n:
            st.warning(
                f"⚠️ You asked for **{int(manual_n)}** calls, but only **{available}** call(s) "
                f"match **{bucket}** in this date range. Showing all {available} available."
            )
        else:
            st.info(f"Showing {int(manual_n)} of {available} matching calls.")
    else:
        selected_df = matched
        st.info(f"Showing all {available} matching calls for **{bucket}**.")

    st.dataframe(selected_df, use_container_width=True, height=300)
    st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # STEP 3 — RUN VAD ANALYSIS
    # ============================================================
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title"><span class="step-badge">3</span>Run Talk-Time Analysis</div>', unsafe_allow_html=True)
    st.markdown('<p class="step-subtitle">Downloads each recording and measures real speaking vs silence time.</p>', unsafe_allow_html=True)

    with st.expander("⚙️ Fine-tune detection accuracy (optional)"):
        st.caption(
            "If Talk Time is coming out too low / Silence too high, move the slider "
            "towards 'Detect more speech'. Default is fine for most calls."
        )
        sensitivity = st.slider(
            "Detection sensitivity",
            min_value=1, max_value=9, value=5,
            help="Lower = detect more speech (fixes low Talk Time). Higher = stricter, only counts confident speech.",
        )
        # Maps slider 1..9 to a VAD confidence threshold 0.15..0.45.
        # Lower threshold = the model needs less confidence to call something
        # "speech", which directly increases Talk Time and reduces Silence
        # Time for calls where speech was previously being missed.
        vad_threshold = round(0.15 + (sensitivity - 1) * (0.45 - 0.15) / 8, 3)
        dead_air_secs = st.number_input(
            "Count a pause as 'Dead Air' only if longer than (sec)",
            min_value=1, value=5, step=1,
        )

    run_vad_clicked = st.button("▶️  Run Analysis & Build Report", type="primary")

    if run_vad_clicked:
        if not col_recording:
            st.error("No recording-URL column found in CDR data — cannot fetch recordings.")
        elif len(selected_df) == 0:
            st.warning("No calls selected — nothing to process.")
        else:
            @st.cache_resource(show_spinner="Loading voice-detection model (first run only)...")
            def load_vad_model():
                torch.hub.set_dir(os.path.expanduser("~/.cache/torch/hub"))
                model, utils = torch.hub.load(
                    "snakers4/silero-vad", "silero_vad", force_reload=False, trust_repo=True
                )
                return model, utils

            model, utils = load_vad_model()
            get_speech_timestamps = utils[0]

            VAD_CFG = {
                "threshold": vad_threshold,
                "min_speech_duration_ms": 100,
                "min_silence_duration_ms": 200,
                "speech_pad_ms": 300,
                "window_size_samples": 512,
                "dead_air_threshold_sec": dead_air_secs,
            }

            def robust_normalize(audio):
                """
                Normalizes loudness using RMS (average energy) instead of peak/
                percentile. Peak-based normalization is fragile: a single loud
                click, beep, or DTMF tone in the recording drags the whole
                clip's volume down when you divide by it, making genuine
                speech quieter than it should be — and the VAD then reads
                that quiet speech as silence, which is exactly what causes
                Talk Time to come out too low / Silence Time too high.
                RMS normalization instead targets the AVERAGE loudness of the
                whole clip, so a short loud spike can't distort it.
                """
                rms = np.sqrt(np.mean(np.square(audio)))
                if rms > 1e-4:  # skip near-total-silence clips, don't amplify noise floor
                    target_rms = 0.1
                    gain = target_rms / rms
                    # cap the gain so we don't blow up extremely quiet recordings
                    # into pure noise
                    gain = min(gain, 20.0)
                    audio = audio * gain
                return np.clip(audio, -1.0, 1.0)

            def load_channel_16k(data, sr, channel_idx=None):
                if data.ndim > 1:
                    chan = data[:, channel_idx] if channel_idx is not None else np.mean(data, axis=1)
                else:
                    chan = data
                chan = robust_normalize(chan.astype(np.float32))
                if sr != 16000:
                    import librosa
                    chan = librosa.resample(chan, orig_sr=sr, target_sr=16000)
                return torch.from_numpy(chan).float()

            def run_vad(audio_tensor):
                return get_speech_timestamps(
                    audio_tensor, model, sampling_rate=16000,
                    threshold=VAD_CFG["threshold"],
                    min_speech_duration_ms=VAD_CFG["min_speech_duration_ms"],
                    min_silence_duration_ms=VAD_CFG["min_silence_duration_ms"],
                    speech_pad_ms=VAD_CFG["speech_pad_ms"],
                    window_size_samples=VAD_CFG["window_size_samples"],
                )

            def merge_intervals(intervals):
                if not intervals:
                    return []
                intervals = sorted(intervals, key=lambda x: x[0])
                merged = [list(intervals[0])]
                for s, e in intervals[1:]:
                    if s <= merged[-1][1]:
                        merged[-1][1] = max(merged[-1][1], e)
                    else:
                        merged.append([s, e])
                return merged

            def compute_metrics(intervals, total_duration):
                """
                Dead Air is NOT separate from Silence Time — it's the portion
                of Silence Time made up of long gaps (> dead_air threshold,
                default 5 sec). Silence Time = total duration - Talk Time.
                """
                if not intervals:
                    return {
                        "talk_time": 0.0,
                        "silence_time": round(total_duration, 2),
                        "dead_air": round(total_duration, 2) if total_duration > VAD_CFG["dead_air_threshold_sec"] else 0.0,
                        "longest_silence": round(total_duration, 2),
                    }
                speech_time, longest_silence, dead_air, prev_end = 0.0, 0.0, 0.0, 0.0
                for s, e in intervals:
                    speech_time += (e - s)
                    silence = max(0.0, s - prev_end)
                    longest_silence = max(longest_silence, silence)
                    if silence > VAD_CFG["dead_air_threshold_sec"]:
                        dead_air += silence
                    prev_end = e
                ending_silence = max(0.0, total_duration - prev_end)
                longest_silence = max(longest_silence, ending_silence)
                if ending_silence > VAD_CFG["dead_air_threshold_sec"]:
                    dead_air += ending_silence
                silence_time = max(0.0, total_duration - speech_time)
                return {
                    "talk_time": round(speech_time, 2),
                    "silence_time": round(silence_time, 2),
                    "dead_air": round(dead_air, 2),
                    "longest_silence": round(longest_silence, 2),
                }

            results = []
            progress = st.progress(0)
            status = st.empty()
            total_rows = len(selected_df)

            with tempfile.TemporaryDirectory() as tmpdir:
                for i, (_, row) in enumerate(selected_df.iterrows()):
                    status.text(f"Processing call {i+1}/{total_rows}...")
                    rec_url = row.get(col_recording)
                    metrics = {
                        "talk_time": None, "silence_time": None,
                        "dead_air": None, "longest_silence": None, "duration": None,
                    }
                    debug_status = "OK"
                    actual_mp3 = None

                    if not rec_url:
                        debug_status = "No recording URL in this row"
                    else:
                        actual_mp3 = resolve_audio_url(rec_url)
                        if not actual_mp3:
                            debug_status = "Could not resolve a direct audio URL from the recording link"
                        else:
                            mp3_path = os.path.join(tmpdir, f"{i}.mp3")
                            wav_path = os.path.join(tmpdir, f"{i}.wav")
                            try:
                                r = requests.get(actual_mp3, timeout=60, stream=True)
                                if r.status_code != 200:
                                    debug_status = f"Download failed: HTTP {r.status_code}"
                                else:
                                    with open(mp3_path, "wb") as f:
                                        for chunk in r.iter_content(8192):
                                            if chunk:
                                                f.write(chunk)
                                    if not os.path.exists(mp3_path) or os.path.getsize(mp3_path) == 0:
                                        debug_status = "Downloaded file is empty"
                                    else:
                                        ff = subprocess.run(
                                            ["ffmpeg", "-y", "-i", mp3_path, "-acodec", "pcm_s16le", wav_path],
                                            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                        )
                                        if ff.returncode != 0 or not os.path.exists(wav_path):
                                            err_tail = ff.stderr.decode(errors="ignore")[-300:]
                                            debug_status = f"FFmpeg conversion failed: {err_tail.strip()}"
                                        else:
                                            data, sr = sf.read(wav_path)
                                            total_duration = len(data) / sr
                                            is_stereo = data.ndim > 1 and data.shape[1] > 1
                                            if is_stereo:
                                                all_intervals = []
                                                for ch in range(data.shape[1]):
                                                    tensor = load_channel_16k(data, sr, channel_idx=ch)
                                                    ts = run_vad(tensor)
                                                    all_intervals.extend([(s["start"] / 16000, s["end"] / 16000) for s in ts])
                                                merged = merge_intervals(all_intervals)
                                            else:
                                                tensor = load_channel_16k(data, sr)
                                                ts = run_vad(tensor)
                                                merged = merge_intervals([(s["start"] / 16000, s["end"] / 16000) for s in ts])
                                            metrics = compute_metrics(merged, total_duration)
                                            metrics["duration"] = round(total_duration, 2)
                                            debug_status = "OK"
                            except requests.exceptions.RequestException as e:
                                debug_status = f"Network/download error: {e}"
                            except Exception as e:
                                debug_status = f"Processing error: {e}"

                    if debug_status != "OK":
                        st.warning(f"Row {i+1} ({row.get(col_agent) if col_agent else ''}): {debug_status}")

                    crm_duration = row.get("_duration_sec")

                    results.append({
                        "Date": row.get(col_date) if col_date else None,
                        "Time": row.get(col_time) if col_time else None,
                        "Agent Name": row.get(col_agent) if col_agent else None,
                        "Call From": row.get(col_phone) if col_phone else None,
                        "Actual MP3": actual_mp3,
                        "Audio Duration(sec)": metrics.get("duration"),
                        "Audio Call Duration": crm_duration,
                        "AI Tools Talk time": metrics.get("talk_time"),
                        "Silence Time": metrics.get("silence_time"),
                        "Dead Air(included in Silence time)": metrics.get("dead_air"),
                        "Longest Silence": metrics.get("longest_silence"),
                        "_debug_status": debug_status,  # internal only — never shown or exported
                    })
                    progress.progress((i + 1) / total_rows)

            status.text("Done ✅")
            final_df = pd.DataFrame(results)
            st.session_state["final_df"] = final_df

            failed_count = (final_df["_debug_status"] != "OK").sum()
            if failed_count > 0:
                st.error(f"⚠️ {failed_count} of {len(final_df)} call(s) failed to process — see warnings above for the reason.")
            else:
                st.success("✅ All calls processed successfully.")

            st.dataframe(final_df.drop(columns=["_debug_status"]), use_container_width=True, height=380)

    st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # STEP 4 — DOWNLOAD FINAL REPORT
    # ============================================================
    if st.session_state.get("final_df") is not None:
        REQUIRED_COLUMNS = [
            "Date", "Time", "Agent Name", "Call From", "Actual MP3",
            "Audio Duration(sec)", "Audio Call Duration", "AI Tools Talk time",
            "Silence Time", "Dead Air(included in Silence time)", "Longest Silence",
        ]
        st.markdown('<div class="step-card">', unsafe_allow_html=True)
        st.markdown('<div class="step-title"><span class="step-badge">4</span>Download Report</div>', unsafe_allow_html=True)
        st.markdown('<p class="step-subtitle">Excel file with exactly the columns you need.</p>', unsafe_allow_html=True)

        export_df = st.session_state["final_df"][REQUIRED_COLUMNS]
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Report")
        buf.seek(0)
        st.download_button(
            "⬇️  Download Excel Report",
            data=buf,
            file_name=f"CallAI_Talk_Time_Report_{client_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("👆 Pick a client and date range above, then click **Fetch Calls** to get started.")

import streamlit as st
import pandas as pd
import requests
import math
import os
import time
import gzip
import shutil
import json
import re
from datetime import datetime, timedelta, timezone
import concurrent.futures

# IST Offset
IST_OFFSET = timedelta(hours=5, minutes=30)
IST = timezone(IST_OFFSET)


def get_ist_now():
    return datetime.now(IST)


# Set page configuration
st.set_page_config(page_title="OTM Positional Scanner", layout="wide")

# Custom CSS for compact layout
st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 1rem !important;
        }
        h1 {
            font-size: 1.8rem !important;
            margin-bottom: 0rem !important;
            white-space: nowrap !important;
        }
        h2 {
            font-size: 1.1rem !important;
            padding-top: 0.2rem !important;
            margin-bottom: 0.1rem !important;
        }
        h3 {
            font-size: 1.0rem !important;
            padding-top: 0.1rem !important;
            margin-bottom: 0.1rem !important;
        }
        /* Prevent graying out during refresh */
        .stApp {
            transition: none !important;
        }
        [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
            opacity: 1 !important;
            transition: none !important;
        }
        /* Hide File Uploader Instructions */
        [data-testid="stFileUploaderDropzone"] div div span {
           display: none !important;
        }
        [data-testid="stFileUploaderDropzone"] div div small {
           display: none !important;
        }
        /* Force Dataframe Font Weight */
        div[data-testid="stDataFrame"] {
            font-weight: 600 !important;
        }
        /* Movers panel ("what changed since last refresh") */
        .movers-panel {
            border: 1px solid #e5e7eb; border-radius: 10px; overflow: hidden; margin-bottom: 14px;
        }
        .movers-head {
            display: flex; align-items: center; justify-content: space-between;
            padding: 8px 14px; background: #f8fafc; border-bottom: 1px solid #e5e7eb;
            font-size: 12.5px; font-weight: 700; color: #334155;
        }
        .movers-empty {
            padding: 10px 14px; font-size: 13px; color: #94a3b8;
        }
        .movers-row {
            display: flex; align-items: center; gap: 10px; padding: 8px 14px;
            border-bottom: 1px solid #f1f5f9; font-size: 13px;
        }
        .movers-row:last-child { border-bottom: none; }
        .movers-arrow-up, .movers-arrow-down {
            width: 22px; height: 22px; border-radius: 6px; display: flex;
            align-items: center; justify-content: center; flex-shrink: 0; font-weight: 700; font-size: 13px;
        }
        .movers-arrow-up { background: #dcfce7; color: #16a34a; }
        .movers-arrow-down { background: #fee2e2; color: #dc2626; }
        .movers-sym { font-weight: 700; color: #0f172a; width: 150px; flex-shrink: 0; }
        .movers-detail { color: #475569; flex: 1; }
        .movers-delta-up { color: #16a34a; font-weight: 700; }
        .movers-delta-down { color: #dc2626; font-weight: 700; }
        /* Movers panel split into CE / PE columns */
        .movers-body { display: flex; }
        .movers-side { flex: 1; min-width: 0; }
        .movers-side:first-child { border-right: 1px solid #e5e7eb; }
        .movers-side-head {
            padding: 6px 14px; font-size: 11px; font-weight: 700; color: #64748b;
            background: #fafbfc; border-bottom: 1px solid #f1f5f9; letter-spacing: 0.03em;
        }
        .movers-side .movers-row { padding: 8px 10px; }
        .movers-side .movers-sym { width: auto; margin-right: 4px; }
    </style>
""", unsafe_allow_html=True)

# Paths for persistent storage
DATA_DIR = 'data'
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

TOKEN_FILE = os.path.join(DATA_DIR, 'token.json')
META_FILE = os.path.join(DATA_DIR, 'meta.json')
LTP_CACHE_FILE = os.path.join(DATA_DIR, 'ltp_cache.json')
TRIGGER_ALERT_FILE = os.path.join(DATA_DIR, 'trigger_alert_state.json')
TRIGGER_TIME_FILE = os.path.join(DATA_DIR, 'trigger_time_state.json')

# Telegram alerts only start firing from this IST time onward (skips the
# noisy pre-open / opening-auction minutes).
ALERT_START_TIME = datetime.strptime("09:30", "%H:%M").time()

# The Monthly / Weekly IV Excel files (produced by the IV Sheet Generator's
# Monthly and Weekly tabs) replace the old NSE Bhavcopy ZIP as the input
# files for this app.
MONTHLY_IV_FILE = os.path.join(DATA_DIR, 'monthly_iv.xlsx')
WEEKLY_IV_FILE = os.path.join(DATA_DIR, 'weekly_iv.xlsx')


def load_meta():
    if os.path.exists(META_FILE):
        try:
            with open(META_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}


def save_meta(key, value):
    try:
        meta = load_meta()
        meta[key] = value
        with open(META_FILE, 'w') as f:
            json.dump(meta, f)
    except:
        pass


def load_ltp_cache():
    if os.path.exists(LTP_CACHE_FILE):
        try:
            with open(LTP_CACHE_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}


def save_ltp_cache(new_data):
    try:
        cache = load_ltp_cache()
        cache.update(new_data)
        with open(LTP_CACHE_FILE, 'w') as f:
            json.dump(cache, f)
    except:
        pass


def extract_expiry_from_filename(filename):
    """
    Extract an expiry date from a filename like 'Monthly IV 25AUG2026.xlsx'
    (produced by the IV Sheet Generator). Returns a normalized (midnight)
    pandas Timestamp, or None if nothing could be parsed.
    """
    # DDMonYYYY e.g. 25AUG2026
    match = re.search(r'(\d{1,2}[A-Za-z]{3}\d{4})', filename)
    if match:
        try:
            return pd.to_datetime(match.group(1), format='%d%b%Y').normalize()
        except Exception:
            pass
    # Fallback: 8-digit date e.g. 20260825
    match8 = re.search(r'(\d{8})', filename)
    if match8:
        d = match8.group(1)
        try:
            return pd.to_datetime(f"{d[:4]}-{d[4:6]}-{d[6:]}").normalize()
        except Exception:
            pass
    return None


def load_token():
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, 'r') as f:
                data = json.load(f)
                if data.get('date') == get_ist_now().strftime('%Y-%m-%d'):
                    return data.get('token', '')
        except:
            pass
    return ''


def save_token(token):
    try:
        data = {
            'date': get_ist_now().strftime('%Y-%m-%d'),
            'token': token
        }
        with open(TOKEN_FILE, 'w') as f:
            json.dump(data, f)
    except:
        pass


# ============================================================
# TELEGRAM TRIGGER-ALERT STATE
#
# Persisted to disk (not just st.session_state) so alert
# de-duplication survives Streamlit Cloud restarts / fragment
# reruns. Resets automatically each new trading day.
# ============================================================
def load_trigger_alert_state():
    if os.path.exists(TRIGGER_ALERT_FILE):
        try:
            with open(TRIGGER_ALERT_FILE, 'r') as f:
                data = json.load(f)
                if data.get('date') == get_ist_now().strftime('%Y-%m-%d'):
                    return set(data.get('keys', []))
        except:
            pass
    return set()


def save_trigger_alert_state(keys):
    try:
        data = {
            'date': get_ist_now().strftime('%Y-%m-%d'),
            'keys': list(keys)
        }
        with open(TRIGGER_ALERT_FILE, 'w') as f:
            json.dump(data, f)
    except:
        pass


# ============================================================
# TRIGGERED-ON TIMESTAMPS (Monthly/Weekly CE/PE tables)
#
# Records the first moment each option's change % actually reaches 100%
# (LTP reaching the Trigger price itself) so the table can show WHEN it
# triggered, not just that it currently reads high.
#
# NOT date-scoped (unlike the Telegram alert-dedup state above) - this is
# keyed by key_suffix + expiry date + instrument, so a "Triggered On"
# stamp holds steady across every refresh AND every day for the entire
# Monthly / Weekly expiry cycle. It only goes away once you upload a new
# IV Excel with a different expiry for that section (the old expiry's keys
# simply stop being looked up).
# ============================================================
def load_trigger_time_state():
    if os.path.exists(TRIGGER_TIME_FILE):
        try:
            with open(TRIGGER_TIME_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}


def save_trigger_time_state(times):
    try:
        data = times
        with open(TRIGGER_TIME_FILE, 'w') as f:
            json.dump(data, f)
    except:
        pass


def clear_trigger_times_for_section(key_suffix):
    """
    Wipes every 'Triggered On' stamp recorded for this section (Monthly or
    Weekly), regardless of expiry. Called the moment a genuinely new IV
    Excel is uploaded for that section, so there's never any ambiguity
    about whether a stamp you're looking at is a fresh same-day trigger or
    a stale leftover from before - a new upload always starts clean.
    """
    times = load_trigger_time_state()
    remaining = {k: v for k, v in times.items() if not k.startswith(f"{key_suffix}:")}
    if len(remaining) != len(times):
        save_trigger_time_state(remaining)


@st.cache_resource
def _get_telegram_session():
    # A reused, persistent connection (kept alive across fragment reruns via
    # st.cache_resource) instead of opening a fresh TCP+TLS handshake to
    # Telegram on every single alert - shaves a meaningful chunk off how
    # long "Send Test" / a live alert takes to actually go out.
    return requests.Session()


def send_telegram_alert(bot_token, chat_id, message):
    if not bot_token or not chat_id:
        return False, "Missing bot token or chat ID"

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "HTML"
    }
    try:
        response = _get_telegram_session().post(url, json=payload, timeout=10)
        if response.status_code == 200:
            return True, None
        return False, f"HTTP {response.status_code}: {response.text[:200]}"
    except Exception as e:
        return False, f"Exception: {e}"


def check_and_alert_triggers(df, key_suffix, telegram_enabled, bot_token, chat_id, threshold_pct=85):
    """
    Sends a Telegram alert the moment an option's change % (LTP vs Trigger)
    reaches threshold_pct (default 85, i.e. 85% of the way to Trigger — not
    only a full 100% cross). Fires once per option per day.

    key_suffix ("Monthly" or "Weekly") picks the bot_token/chat_id this
    call was given, so Monthly and Weekly alerts are fully independent —
    each section sends its own message to its own configured bot/chat,
    never combined with the other section's rows.

    Alerts are suppressed before ALERT_START_TIME (09:30 IST) so the noisy
    opening minutes don't spam Telegram — but crucially, anything already
    at/above threshold_pct *before* 09:30 is silently marked as "seen"
    (never actually messaged) so it doesn't get dumped as a fake "just
    crossed" alert the instant the clock ticks past 09:30. Only strikes
    that genuinely cross the threshold AFTER 09:30 ever produce a message.
    """
    if not telegram_enabled:
        return
    if df.empty:
        return
    if 'instrument_key' not in df.columns:
        return

    ist_now = get_ist_now()
    alerted = load_trigger_alert_state()

    if ist_now.time() < ALERT_START_TIME:
        # Pre-market / opening-auction window: silently record anything
        # already at/above threshold so it's excluded once alerting turns
        # on at 09:30, instead of firing in a batch right at that moment.
        pre_seen = set()
        for _, row in df.iterrows():
            inst_key = row.get('instrument_key')
            if not inst_key or pd.isna(inst_key):
                continue
            try:
                change_pct = float(row.get('change %', 0.0))
            except:
                continue
            if change_pct >= threshold_pct:
                pre_seen.add(f"{key_suffix}:{inst_key}")
        if pre_seen - alerted:
            save_trigger_alert_state(alerted | pre_seen)
        return

    newly_triggered = []
    for _, row in df.iterrows():
        inst_key = row.get('instrument_key')
        if not inst_key or pd.isna(inst_key):
            continue

        alert_id = f"{key_suffix}:{inst_key}"
        try:
            change_pct = float(row.get('change %', 0.0))
        except:
            continue

        if change_pct >= threshold_pct and alert_id not in alerted:
            newly_triggered.append(row)
            alerted.add(alert_id)

    if not newly_triggered:
        return

    trigger_time_label = ist_now.strftime('%d-%b-%Y %H:%M:%S')
    header = f"🚀 <b>{threshold_pct:.0f}% Hit — {key_suffix}</b> · {trigger_time_label} IST"
    option_lines = [
        f"{row['Symbol']} {row['StrikePrice']:.0f} {row['OptionType']} | LTP {row['ltp']:.2f} | Trig {row['Trigger']:.2f}"
        for row in newly_triggered
    ]
    message = header + "\n" + "\n".join(option_lines)

    success, error = send_telegram_alert(bot_token, chat_id, message)
    if success:
        save_trigger_alert_state(alerted)
        # st.toast (not st.sidebar.success/warning): this function runs inside an
        # @st.fragment (show_monthly/show_weekly). Writing to st.sidebar - a
        # container outside the fragment's own tree - from inside a fragment
        # raises StreamlitAPIException ("container was not written to during
        # the initial run") and aborts the fragment mid-run, which is why the
        # option tables were disappearing whenever an alert fired. st.toast()
        # is a floating overlay that doesn't need a reserved container, so it's
        # safe to call from any fragment.
        st.toast(f"📨 Telegram alert sent for {len(newly_triggered)} trigger cross(es) on {key_suffix}.", icon="✅")
    else:
        st.toast(f"⚠️ Telegram alert failed ({key_suffix}): {error}", icon="⚠️")


# Constant for NSE JSON
NSE_JSON_PATH = 'NSE.json'


@st.cache_data
def load_nse_json():
    if os.path.exists(NSE_JSON_PATH):
        try:
            df = pd.read_json(NSE_JSON_PATH)
            if 'segment' in df.columns:
                df = df[df['segment'] == 'NSE_FO']
            df['expiry_dt'] = pd.to_datetime(df['expiry'], unit='ms').dt.normalize()
            return df
        except Exception as e:
            st.error(f"Error loading NSE.json: {e}")
            return pd.DataFrame()
    else:
        st.error(f"NSE.json not found at {NSE_JSON_PATH}")
        return pd.DataFrame()


def process_iv_excel(excel_path, df_json, expiry_date):
    """
    Reads the Monthly IV Excel (from the IV Sheet Generator) and builds the
    OTM universe this scanner tracks:
        Upper Strike -> CE (Call) side
        Lower Strike -> PE (Put) side

    Trigger = Close price x 2
    TGT     = Trigger x 2
    SL      = Trigger / 2
    """
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"Failed to read Monthly IV Excel: {e}")
        return pd.DataFrame()

    required_cols = ['NAME', 'UPPER STRIKE PRICE', 'UPPER STRIKE CLOSE PRICE',
                      'LOWER STRIKE PRICE', 'LOWER STRIKE CLOSE PRICE']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Uploaded Excel missing required columns: {missing}")
        return pd.DataFrame()

    df = df.dropna(subset=['NAME'])

    # Upper Strike -> CE leg
    upper = df[['NAME', 'UPPER STRIKE PRICE', 'UPPER STRIKE CLOSE PRICE']].copy()
    upper = upper.rename(columns={
        'NAME': 'Symbol',
        'UPPER STRIKE PRICE': 'StrikePrice',
        'UPPER STRIKE CLOSE PRICE': 'Close'
    })
    upper['OptionType'] = 'CE'

    # Lower Strike -> PE leg
    lower = df[['NAME', 'LOWER STRIKE PRICE', 'LOWER STRIKE CLOSE PRICE']].copy()
    lower = lower.rename(columns={
        'NAME': 'Symbol',
        'LOWER STRIKE PRICE': 'StrikePrice',
        'LOWER STRIKE CLOSE PRICE': 'Close'
    })
    lower['OptionType'] = 'PE'

    combined = pd.concat([upper, lower], ignore_index=True)
    combined = combined.dropna(subset=['StrikePrice', 'Close'])
    combined = combined[combined['StrikePrice'] > 0]

    if combined.empty:
        return pd.DataFrame()

    combined['ExpiryDate'] = expiry_date
    # Round to avoid float-precision mismatches on the merge key below
    combined['StrikePrice'] = combined['StrikePrice'].round(2)

    if df_json is None or df_json.empty:
        st.warning("NSE.json not loaded — cannot map instrument keys / fetch live LTP.")
        combined['instrument_key'] = None
    else:
        df_json = df_json.copy()
        df_json['strike_price'] = df_json['strike_price'].astype(float).round(2)

        merged = pd.merge(
            combined,
            df_json,
            left_on=['Symbol', 'StrikePrice', 'OptionType', 'ExpiryDate'],
            right_on=['underlying_symbol', 'strike_price', 'instrument_type', 'expiry_dt'],
            how='left'
        )

        if merged['instrument_key'].isna().all() and not merged.empty:
            st.error(
                "Data mismatch: Could not find any of these strikes in NSE.json. "
                "Check that the Expiry Date set in the sidebar matches this Monthly IV file, "
                "or update NSE.json."
            )

        combined = merged[['Symbol', 'StrikePrice', 'OptionType', 'Close', 'instrument_key']]

    # Trigger / Target calculation (User Rule)
    combined['Trigger'] = combined['Close'] * 2
    combined['TGT'] = combined['Trigger'] * 2
    combined['SL'] = combined['Trigger'] / 2

    return combined.reset_index(drop=True)


def fetch_ltp(instrument_keys, token):
    if not token:
        return {}

    url = "https://api.upstox.com/v3/market-quote/ltp"
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {token}'
    }

    batch_size = 50
    ltp_map = {}
    batches = [instrument_keys[i:i + batch_size] for i in range(0, len(instrument_keys), batch_size)]

    def fetch_batch(batch):
        params = {'instrument_key': ','.join(batch)}
        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 'success':
                    quotes = data.get('data', {})
                    result = {}
                    for key, details in quotes.items():
                        inst_token = details.get('instrument_token')
                        last_price = details.get('last_price')
                        if inst_token is not None:
                            result[inst_token] = last_price
                    return result
        except Exception:
            pass
        return {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fetch_batch, batch) for batch in batches]
        for future in concurrent.futures.as_completed(futures):
            try:
                batch_result = future.result()
                if batch_result:
                    ltp_map.update(batch_result)
            except Exception:
                pass

    return ltp_map


# An option counts as "triggered" once its change % (LTP vs Trigger) first
# reaches this — i.e. LTP has actually reached the Trigger price itself.
TRIGGERED_AT_PCT = 100.0


def attach_trigger_times(df, key_suffix, expiry_date):
    """
    Stamps a 'Triggered On' column onto df: the date & time the option's
    change % FIRST reached TRIGGERED_AT_PCT (100%, i.e. LTP actually hit
    the Trigger price). Once recorded it NEVER changes again — not for the
    rest of the day, and not on later days either — for as long as this
    Monthly/Weekly section keeps the same expiry_date. It only resets once
    you upload a new IV Excel with a different expiry (a new expiry_date
    means a brand-new set of keys, so the old stamps simply stop applying).

    Persisted to disk so it survives refreshes/reruns. Blank ('—') until
    the option actually triggers.

    Format: "<day-of-month>&<hour>.<minute>" in 12-hour clock, e.g. a
    trigger at 13:35 IST on the 24th shows as "24&1.35" (the day/time it
    FIRST happened, not today's date).
    """
    times = load_trigger_time_state()
    changed = False

    now_dt = get_ist_now()
    hour_12 = now_dt.hour % 12
    if hour_12 == 0:
        hour_12 = 12
    now_label = f"{now_dt.day}&{hour_12}.{now_dt.minute:02d}"

    expiry_key = expiry_date.strftime('%Y-%m-%d') if expiry_date is not None else 'noexpiry'

    labels = []
    for _, row in df.iterrows():
        inst_key = row.get('instrument_key')
        if not inst_key or pd.isna(inst_key):
            labels.append('—')
            continue

        key = f"{key_suffix}:{expiry_key}:{inst_key}"
        if key in times:
            labels.append(times[key])
            continue

        try:
            change_pct = float(row.get('change %', 0.0))
        except:
            change_pct = 0.0

        if change_pct >= TRIGGERED_AT_PCT:
            times[key] = now_label
            changed = True
            labels.append(now_label)
        else:
            labels.append('—')

    if changed:
        save_trigger_time_state(times)

    df = df.copy()
    df['Triggered On'] = labels
    return df


# Minimum move (in change % percentage points) for a row to be shown in
# the "What Changed Since Last Refresh" movers panel — keeps the panel to
# genuine movers instead of every tiny tick.
MOVERS_MIN_DELTA = 1.0
MOVERS_MAX_ROWS = 8
# Movers panel only shows options whose CURRENT change % (LTP vs Trigger)
# is within this band — too far below 75% isn't worth watching yet, and
# above 125% is already well past Trigger.
MOVERS_MIN_PCT = 75.0
MOVERS_MAX_PCT = 125.0


def compute_and_render_movers(df, key_suffix):
    """
    Compares the current change % for each instrument against the snapshot
    taken on the previous refresh (stored in st.session_state, so it works
    whether the "previous refresh" was an auto-refresh fragment rerun or a
    plain manual rerun). Only shows rows that moved UP by >= MOVERS_MIN_DELTA
    percentage points AND whose current change % is between MOVERS_MIN_PCT
    and MOVERS_MAX_PCT, split into separate CE / PE columns, biggest movers
    first.
    """
    snapshot_key = f'movers_prev_{key_suffix}'
    snapshot_time_key = f'movers_prev_time_{key_suffix}'

    prev_snapshot = st.session_state.get(snapshot_key, {})
    prev_time = st.session_state.get(snapshot_time_key)

    current_snapshot = {}
    ce_movers = []
    pe_movers = []

    for _, row in df.iterrows():
        inst_key = row.get('instrument_key')
        if not inst_key or pd.isna(inst_key):
            continue

        try:
            cur_change = float(row['change %'])
        except:
            continue

        current_snapshot[inst_key] = cur_change

        if inst_key in prev_snapshot:
            delta = cur_change - prev_snapshot[inst_key]
            # Only positive (upward) moves, and only within the 75%-125% band.
            if delta >= MOVERS_MIN_DELTA and MOVERS_MIN_PCT <= cur_change <= MOVERS_MAX_PCT:
                mover = {
                    'Symbol': row['Symbol'],
                    'StrikePrice': row['StrikePrice'],
                    'Prev': prev_snapshot[inst_key],
                    'Now': cur_change,
                    'Delta': delta,
                }
                if row['OptionType'] == 'CE':
                    ce_movers.append(mover)
                else:
                    pe_movers.append(mover)

    # Save this run's snapshot for the *next* refresh to compare against.
    st.session_state[snapshot_key] = current_snapshot
    st.session_state[snapshot_time_key] = get_ist_now()

    ce_movers.sort(key=lambda m: m['Delta'], reverse=True)
    pe_movers.sort(key=lambda m: m['Delta'], reverse=True)
    ce_movers = ce_movers[:MOVERS_MAX_ROWS]
    pe_movers = pe_movers[:MOVERS_MAX_ROWS]

    since_label = prev_time.strftime('%H:%M:%S') if prev_time else "—"

    def render_side(movers, label):
        side_html = [f'<div class="movers-side"><div class="movers-side-head">{label} &middot; {len(movers)} mover(s)</div>']
        if prev_time is None:
            side_html.append('<div class="movers-empty">Collecting baseline…</div>')
        elif not movers:
            side_html.append('<div class="movers-empty">No qualifying mover (75%-125% band).</div>')
        else:
            for m in movers:
                side_html.append(
                    '<div class="movers-row">'
                    '<div class="movers-arrow-up">▲</div>'
                    f'<div class="movers-sym">{m["Symbol"]} {m["StrikePrice"]:.0f}</div>'
                    f'<div class="movers-detail">{m["Prev"]:.2f}% &rarr; {m["Now"]:.2f}%</div>'
                    f'<div class="movers-delta-up">+{m["Delta"]:.2f}%</div>'
                    '</div>'
                )
        side_html.append('</div>')
        return "".join(side_html)

    html = ['<div class="movers-panel">']
    html.append(
        f'<div class="movers-head"><span>WHAT CHANGED SINCE LAST REFRESH &middot; {since_label} IST</span>'
        f'<span>{len(ce_movers)} CE &middot; {len(pe_movers)} PE</span></div>'
    )
    html.append('<div class="movers-body">')
    html.append(render_side(ce_movers, "CE"))
    html.append(render_side(pe_movers, "PE"))
    html.append('</div>')
    html.append('</div>')

    st.markdown("".join(html), unsafe_allow_html=True)


def display_option_chain(df, access_token, key_suffix, expiry_date=None, telegram_enabled=False, telegram_bot_token="", telegram_chat_id="", alert_threshold_pct=85):
    st.caption(f"Last Updated: {get_ist_now().strftime('%H:%M:%S')} IST")

    if df.empty:
        st.info("No data to display. Please upload a valid Monthly IV Excel in the sidebar.")
        return

    # Fetch LTP if token provided
    if access_token:
        all_keys = df['instrument_key'].dropna().unique().tolist()

        ist_now = get_ist_now()
        current_time = ist_now.time()
        start_time = datetime.strptime("09:00", "%H:%M").time()
        end_time = datetime.strptime("15:40", "%H:%M").time()
        is_market_hours = start_time <= current_time <= end_time

        ltp_cache = load_ltp_cache()
        missing_keys = [k for k in all_keys if k not in ltp_cache]

        force_refresh = st.session_state.get('force_refresh_ltp', False)

        should_fetch = False
        if is_market_hours:
            should_fetch = True
        elif force_refresh:
            should_fetch = True
            st.session_state['force_refresh_ltp'] = False
        elif missing_keys:
            should_fetch = True

        if should_fetch:
            keys_to_fetch = all_keys if is_market_hours else missing_keys
            fetched_data = fetch_ltp(keys_to_fetch, access_token)
            if fetched_data:
                save_ltp_cache(fetched_data)
                ltp_cache = load_ltp_cache()

        ltp_data = {k: ltp_cache.get(k, 0.0) for k in all_keys}
        df['ltp'] = df['instrument_key'].map(ltp_data).fillna(0.0)
    else:
        df['ltp'] = 0.0
        st.warning("Enter Access Token in sidebar to see live LTP.")

    # Calculate Change % (LTP vs Trigger)
    def calculate_numeric_change(row):
        try:
            trigger = row['Trigger']
            ltp = row['ltp']
            if trigger > 0 and ltp > 0:
                return (ltp / trigger * 100)
            return 0.0
        except:
            return 0.0

    df['change %'] = df.apply(calculate_numeric_change, axis=1)

    # Drop options whose Trigger price is below ₹3 — too cheap to be a
    # meaningful/tradeable OTM candidate, so keep them out of the table,
    # the movers panel, and the Telegram alerts entirely.
    df = df[df['Trigger'] >= 3].copy()
    if df.empty:
        st.info("No rows with Trigger price ≥ ₹3.")
        return

    # --- Telegram Trigger Alerts (>= alert_threshold_pct, only from 09:30 IST) ---
    # bot_token/chat_id passed in are section-specific (Monthly vs Weekly),
    # so each section's alerts go to its own configured Telegram bot/chat.
    check_and_alert_triggers(df, key_suffix, telegram_enabled, telegram_bot_token, telegram_chat_id, alert_threshold_pct)

    # --- What Changed Since Last Refresh ---
    compute_and_render_movers(df, key_suffix)

    # Split Upper Strike (CE) / Lower Strike (PE)
    calls_df = df[df['OptionType'] == 'CE'].copy()
    puts_df = df[df['OptionType'] == 'PE'].copy()

    calls_df = calls_df.sort_values(by='change %', ascending=False)
    puts_df = puts_df.sort_values(by='change %', ascending=False)

    display_cols = ['Symbol', 'StrikePrice', 'ltp', 'Trigger', 'change %', 'TGT', 'SL']

    def color_change(val):
        # Fixed two-tier coloring (no graduated/ascending scale):
        # >=100 -> dark green, 90-99 -> light green, below 90 -> no color.
        if not isinstance(val, (int, float)):
            return ''
        if val >= 100:
            return 'background-color: darkgreen; color: white; font-weight: 700'
        elif val >= 90:
            return 'background-color: lightgreen; color: black; font-weight: 700'
        return ''

    format_dict = {
        'change %': '{:.2f}%',
        'Trigger': '{:.2f}',
        'TGT': '{:.2f}',
        'SL': '{:.2f}',
        'ltp': '{:.2f}',
        'StrikePrice': '{:.2f}'
    }

    def render_table(data_df):
        return (
            data_df[display_cols].style
            .map(color_change, subset=['change %'])
            .set_properties(subset=['TGT'], **{'color': '#1a73e8'})
            .set_properties(subset=['SL'], **{'background-color': '#fdecea', 'color': '#c0392b'})
            .format(format_dict)
            .set_properties(**{'font-weight': '600', 'text-align': 'center', 'font-size': '16px'})
        )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Upper Strike (CE)")
        calls_df = calls_df.sort_values(by='change %', ascending=False)
        st.dataframe(
            render_table(calls_df),
            hide_index=True,
            width='stretch',
            height=1800,
        )
    with col2:
        st.subheader("Lower Strike (PE)")
        puts_df = puts_df.sort_values(by='change %', ascending=False)
        st.dataframe(
            render_table(puts_df),
            hide_index=True,
            width='stretch',
            height=1800,
        )


# --- Configuration Logic (Before Sidebar) ---
# Wrapped in try/except: st.secrets raises if no secrets.toml exists at all
# (e.g. running locally without one configured) - default to Admin view in that case.
try:
    is_client_view = "UPSTOX_ACCESS_TOKEN" in st.secrets and st.secrets["UPSTOX_ACCESS_TOKEN"].strip() != ""
except Exception:
    is_client_view = False

if is_client_view:
    access_token = st.secrets["UPSTOX_ACCESS_TOKEN"]
    st.markdown("""
    <style>
        [data-testid="stSidebar"] {display: none;}
    </style>
    """, unsafe_allow_html=True)
    auto_refresh = True
    refresh_interval = 15

    # Monthly and Weekly each get their own bot/chat, falling back to the
    # shared TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID / ALERT_THRESHOLD_PCT
    # secrets if a section-specific one isn't set.
    monthly_telegram_bot_token = st.secrets.get("TELEGRAM_BOT_TOKEN_MONTHLY", st.secrets.get("TELEGRAM_BOT_TOKEN", ""))
    monthly_telegram_chat_id = st.secrets.get("TELEGRAM_CHAT_ID_MONTHLY", st.secrets.get("TELEGRAM_CHAT_ID", ""))
    monthly_telegram_enabled = bool(monthly_telegram_bot_token and monthly_telegram_chat_id)
    monthly_alert_threshold_pct = float(st.secrets.get("ALERT_THRESHOLD_PCT_MONTHLY", st.secrets.get("ALERT_THRESHOLD_PCT", 85.0)))

    weekly_telegram_bot_token = st.secrets.get("TELEGRAM_BOT_TOKEN_WEEKLY", st.secrets.get("TELEGRAM_BOT_TOKEN", ""))
    weekly_telegram_chat_id = st.secrets.get("TELEGRAM_CHAT_ID_WEEKLY", st.secrets.get("TELEGRAM_CHAT_ID", ""))
    weekly_telegram_enabled = bool(weekly_telegram_bot_token and weekly_telegram_chat_id)
    weekly_alert_threshold_pct = float(st.secrets.get("ALERT_THRESHOLD_PCT_WEEKLY", st.secrets.get("ALERT_THRESHOLD_PCT", 85.0)))
else:
    with st.sidebar:
        st.header("Configuration")
        saved_token = load_token()
        access_token = st.text_input("Upstox Access Token", value=saved_token, type="password")
        if access_token and access_token != saved_token:
            save_token(access_token)

        st.markdown("---")
        st.header("Telegram Alerts")
        st.caption("🕤 Alerts are silent before 09:30 AM IST, then active for the rest of the day.")

        def render_telegram_section(section_prefix, section_label):
            """
            Renders one section's (Monthly or Weekly) independent Telegram
            config: its own enable checkbox, threshold, bot token, chat ID
            and test button - so Monthly and Weekly alerts can go to
            completely different Telegram bots/chats and never get mixed
            into the same message.

            Returns (enabled, bot_token, chat_id, threshold_pct).
            """
            st.subheader(f"{section_label}")
            enabled = st.checkbox(
                f"Enable {section_label} Trigger Alerts",
                value=st.session_state.get(f'{section_prefix}_telegram_enabled', False),
                key=f'{section_prefix}_telegram_enabled',
                help=f"Sends a Telegram message once a {section_label} option's change % (LTP vs Trigger) reaches the threshold below."
            )
            threshold_pct = st.number_input(
                "Alert Threshold (%)",
                min_value=1.0,
                max_value=300.0,
                value=st.session_state.get(f'{section_prefix}_alert_threshold_pct', 85.0),
                step=1.0,
                key=f'{section_prefix}_alert_threshold_pct',
                help="Telegram alert fires once change % reaches this value (default 85, i.e. before the full 100% Trigger cross)."
            )
            bot_token = st.text_input(
                "Bot Token",
                type="password",
                value=st.session_state.get(f'{section_prefix}_telegram_bot_token', ''),
                key=f'{section_prefix}_telegram_bot_token',
                help="Create a bot via @BotFather on Telegram to get this token."
            )
            chat_id = st.text_input(
                "Chat ID",
                value=st.session_state.get(f'{section_prefix}_telegram_chat_id', ''),
                key=f'{section_prefix}_telegram_chat_id',
                help="Your personal or group chat ID. Message @userinfobot to find yours."
            )

            tg_col1, tg_col2 = st.columns(2)
            test_clicked = tg_col1.button("Send Test", use_container_width=True, key=f'{section_prefix}_test_btn')
            reset_clicked = tg_col2.button("Reset Alerts", use_container_width=True, key=f'{section_prefix}_reset_btn')

            if reset_clicked:
                # Alert-dedup keys are namespaced "Monthly:..."/"Weekly:...",
                # so only this section's already-triggered options are
                # cleared - the other section's alert history is untouched.
                alerted = load_trigger_alert_state()
                remaining = {k for k in alerted if not k.startswith(f"{section_label}:")}
                save_trigger_alert_state(remaining)
                st.success(f"{section_label} alert state cleared — already-triggered options will alert again.")

            if test_clicked:
                success, error = send_telegram_alert(
                    bot_token,
                    chat_id,
                    f"✅ Test alert from OTM Positional Scanner ({section_label}) — Telegram is wired up correctly."
                )
                if success:
                    st.success("Test message sent — check Telegram.")
                else:
                    st.error(f"Test message failed: {error}")

            return enabled, bot_token, chat_id, threshold_pct

        monthly_telegram_enabled, monthly_telegram_bot_token, monthly_telegram_chat_id, monthly_alert_threshold_pct = render_telegram_section('monthly', 'Monthly')
        st.markdown("---")
        weekly_telegram_enabled, weekly_telegram_bot_token, weekly_telegram_chat_id, weekly_alert_threshold_pct = render_telegram_section('weekly', 'Weekly')

        st.markdown("---")
        st.header("Data Management")

        if st.button("⚡ Refresh LTP Now", use_container_width=True):
            st.session_state['force_refresh_ltp'] = True
            st.rerun()

        # NSE JSON Uploader
        st.subheader("NSE Instrument JSON")
        if st.button("🔄 Download Latest"):
            try:
                with st.spinner("Downloading latest NSE.json from Upstox..."):
                    url = "https://assets.upstox.com/market-quote/instruments/exchange/NSE.json.gz"
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
                    }
                    response = requests.get(url, headers=headers, stream=True)
                    if response.status_code == 200:
                        with open(NSE_JSON_PATH, "wb") as f_out:
                            with gzip.GzipFile(fileobj=response.raw) as f_in:
                                shutil.copyfileobj(f_in, f_out)
                        st.cache_data.clear()
                        st.success("Updated successfully!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(f"Failed to download. Status: {response.status_code}")
            except Exception as e:
                st.error(f"Error: {e}")

        st.markdown("---")

        def render_iv_upload_section(section_key, file_path, meta_file_key, meta_expiry_key, label, example_name, trigger_key_suffix):
            """
            Renders one Upload + Confirm Expiry block (used for both the
            Monthly IV Excel and the Weekly IV Excel sections below).

            section_key keeps each block's widget keys independent so the
            two sections never clash with each other. trigger_key_suffix is
            the "Monthly"/"Weekly" namespace used by attach_trigger_times,
            so a fresh upload here can wipe that section's Triggered On
            history.
            """
            st.subheader(label)
            up_iv = st.file_uploader(
                f"Upload {label}",
                type=['xlsx'],
                key=f'{section_key}_iv_up',
                help=f"The output file from the IV Sheet Generator (e.g. '{example_name}')."
            )

            if up_iv is not None:
                # st.file_uploader keeps returning the SAME UploadedFile on
                # every single rerun for as long as it sits in the widget -
                # not just on the rerun where it was chosen. Several actions
                # elsewhere in this app trigger a full-page st.rerun(), and without
                # this guard every one of those reruns would re-detect the
                # expiry from the SAME original filename and stomp any
                # manual correction made afterwards in "Confirm Expiry
                # Date". file_id (or name+size as a fallback) uniquely
                # identifies one upload event, so this block now only runs
                # once per GENUINE new upload.
                upload_marker = getattr(up_iv, 'file_id', None) or f"{up_iv.name}:{up_iv.size}"
                processed_key = f'{section_key}_last_processed_upload'

                if st.session_state.get(processed_key) != upload_marker:
                    with open(file_path, "wb") as f:
                        f.write(up_iv.getvalue())
                    save_meta(meta_file_key, up_iv.name)

                    # A genuinely new file for this section always starts
                    # Triggered On fresh - no ambiguity about whether a
                    # stamp is a real same-day trigger or a stale leftover.
                    clear_trigger_times_for_section(trigger_key_suffix)

                    detected_expiry = extract_expiry_from_filename(up_iv.name)
                    if detected_expiry is not None:
                        save_meta(meta_expiry_key, detected_expiry.strftime('%Y-%m-%d'))
                        # The "Confirm Expiry Date" widget below has a
                        # persistent key, so Streamlit would otherwise keep
                        # showing whatever date was last picked for the
                        # PREVIOUS file. Clearing its session state forces
                        # it to re-initialize from the freshly saved meta
                        # value instead.
                        st.session_state.pop(f'{section_key}_expiry_input', None)
                        st.success(f"Uploaded {up_iv.name} — Expiry detected: {detected_expiry.strftime('%d-%b-%Y')}")
                    else:
                        st.warning(f"Uploaded {up_iv.name} — could not auto-detect expiry from filename. Please confirm it below.")

                    st.session_state[processed_key] = upload_marker

            meta = load_meta()
            if os.path.exists(file_path):
                st.caption(f"📄 File: {meta.get(meta_file_key, os.path.basename(file_path))}")

                stored_expiry_str = meta.get(meta_expiry_key)
                try:
                    default_expiry_date = datetime.strptime(stored_expiry_str, '%Y-%m-%d').date() if stored_expiry_str else get_ist_now().date()
                except Exception:
                    default_expiry_date = get_ist_now().date()

                manual_expiry = st.date_input(
                    "Confirm Expiry Date",
                    value=default_expiry_date,
                    key=f'{section_key}_expiry_input',
                    help="Must match the option expiry exactly, so it can be matched against NSE.json."
                )
                save_meta(meta_expiry_key, manual_expiry.strftime('%Y-%m-%d'))

        # Monthly IV Excel Uploader (replaces the old Bhavcopy ZIP upload)
        render_iv_upload_section(
            section_key='monthly',
            file_path=MONTHLY_IV_FILE,
            meta_file_key='MonthlyIVFileName',
            meta_expiry_key='MonthlyIVExpiry',
            label='Monthly IV Excel',
            example_name='Monthly IV 25AUG2026.xlsx',
            trigger_key_suffix='Monthly'
        )

        st.markdown("---")

        # Weekly IV Excel Uploader
        render_iv_upload_section(
            section_key='weekly',
            file_path=WEEKLY_IV_FILE,
            meta_file_key='WeeklyIVFileName',
            meta_expiry_key='WeeklyIVExpiry',
            label='Weekly IV Excel',
            example_name='Weekly IV 29AUG2026.xlsx',
            trigger_key_suffix='Weekly'
        )

        st.markdown("---")
        st.header("Auto Refresh")
        auto_refresh = st.checkbox("Enable Auto-Refresh", value=False)
        refresh_interval = st.slider("Refresh Interval (seconds)", min_value=5, max_value=60, value=15)

# --- Main Page ---
st.title("OTM Positional Scanner")

nse_json_df = load_nse_json()


def get_target_expiry(meta_expiry_key):
    meta = load_meta()
    expiry_str = meta.get(meta_expiry_key)
    if expiry_str:
        try:
            return pd.to_datetime(expiry_str).normalize()
        except Exception:
            return None
    return None


if not nse_json_df.empty:
    tab_monthly, tab_weekly = st.tabs(["📅 Monthly", "🗓️ Weekly"])

    with tab_monthly:
        st.header("Monthly Options")
        target_expiry_m = get_target_expiry('MonthlyIVExpiry')

        if os.path.exists(MONTHLY_IV_FILE) and target_expiry_m is not None:
            st.info(f"📅 Displaying Expiry: **{target_expiry_m.strftime('%d-%b-%Y')}**")

            run_every = refresh_interval if auto_refresh else None

            @st.fragment(run_every=run_every)
            def show_monthly():
                df_m = process_iv_excel(MONTHLY_IV_FILE, nse_json_df, target_expiry_m)
                display_option_chain(df_m, access_token, "Monthly", expiry_date=target_expiry_m, telegram_enabled=monthly_telegram_enabled, telegram_bot_token=monthly_telegram_bot_token, telegram_chat_id=monthly_telegram_chat_id, alert_threshold_pct=monthly_alert_threshold_pct)

            show_monthly()
        else:
            st.warning("Monthly IV Excel file not found. Please upload it in the sidebar.")

    with tab_weekly:
        st.header("Weekly Options")
        target_expiry_w = get_target_expiry('WeeklyIVExpiry')

        if os.path.exists(WEEKLY_IV_FILE) and target_expiry_w is not None:
            st.info(f"📅 Displaying Expiry: **{target_expiry_w.strftime('%d-%b-%Y')}**")

            run_every = refresh_interval if auto_refresh else None

            @st.fragment(run_every=run_every)
            def show_weekly():
                df_w = process_iv_excel(WEEKLY_IV_FILE, nse_json_df, target_expiry_w)
                display_option_chain(df_w, access_token, "Weekly", expiry_date=target_expiry_w, telegram_enabled=weekly_telegram_enabled, telegram_bot_token=weekly_telegram_bot_token, telegram_chat_id=weekly_telegram_chat_id, alert_threshold_pct=weekly_alert_threshold_pct)

            show_weekly()
        else:
            st.warning("Weekly IV Excel file not found. Please upload it in the sidebar.")
else:
    st.error("Critical Error: NSE.json could not be loaded.")

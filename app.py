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

# The Monthly IV Excel (produced by the IV Sheet Generator) replaces the
# old NSE Bhavcopy ZIP as the input file for this app.
MONTHLY_IV_FILE = os.path.join(DATA_DIR, 'monthly_iv.xlsx')

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
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            return True, None
        return False, f"HTTP {response.status_code}: {response.text[:200]}"
    except Exception as e:
        return False, f"Exception: {e}"


def check_and_alert_triggers(df, key_suffix, telegram_enabled, bot_token, chat_id):
    """
    Sends a Telegram alert the moment an option's LTP crosses its
    Trigger price (change % >= 100). Fires once per option per day.
    """
    if not telegram_enabled:
        return

    if df.empty:
        return

    if 'instrument_key' not in df.columns:
        return

    alerted = load_trigger_alert_state()
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

        if change_pct >= 100 and alert_id not in alerted:
            newly_triggered.append(row)
            alerted.add(alert_id)

    if not newly_triggered:
        return

    message_lines = [f"🚀 <b>Trigger Crossed — {key_suffix}</b>"]
    for row in newly_triggered:
        tgt_hit = row.get('TGT HIT', '-')
        message_lines.append(
            f"\n<b>{row['Symbol']} {row['StrikePrice']:.0f} {row['OptionType']}</b>\n"
            f"LTP: {row['ltp']:.2f}  ›  Trigger: {row['Trigger']:.2f}  ›  TGT: {row['TGT']:.2f}\n"
            f"Change: {row['change %']:.2f}%   |   TGT Status: {tgt_hit}"
        )

    message = "\n".join(message_lines)
    success, error = send_telegram_alert(bot_token, chat_id, message)

    if success:
        save_trigger_alert_state(alerted)
        st.sidebar.success(f"Telegram alert sent for {len(newly_triggered)} trigger cross(es).")
    else:
        st.sidebar.warning(f"Telegram alert failed: {error}")


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


def display_option_chain(df, access_token, key_suffix, telegram_enabled=False, telegram_bot_token="", telegram_chat_id=""):
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

    # TGT HIT column
    def calculate_tgt_hit(row):
        try:
            if row['ltp'] > 0 and row['ltp'] >= row['TGT']:
                return 'TGT HIT'
            return '-'
        except:
            return '-'

    df['TGT HIT'] = df.apply(calculate_tgt_hit, axis=1)

    # --- Telegram Trigger Alerts ---
    check_and_alert_triggers(df, key_suffix, telegram_enabled, telegram_bot_token, telegram_chat_id)

    # Split Upper Strike (CE) / Lower Strike (PE)
    calls_df = df[df['OptionType'] == 'CE'].copy()
    puts_df = df[df['OptionType'] == 'PE'].copy()

    calls_df = calls_df.sort_values(by='change %', ascending=False)
    puts_df = puts_df.sort_values(by='change %', ascending=False)

    display_cols = ['Symbol', 'StrikePrice', 'Trigger', 'TGT', 'ltp', 'change %', 'TGT HIT']

    def color_change(val):
        if isinstance(val, (int, float)):
            if val >= 100:
                return 'background-color: darkgreen; color: white'
            elif val >= 90:
                return 'background-color: lightgreen; color: black'
        return ''

    def color_tgt_hit(val):
        if val == 'TGT HIT':
            return 'background-color: darkgreen; color: white'
        return ''

    format_dict = {
        'change %': '{:.2f}%',
        'Trigger': '{:.2f}',
        'TGT': '{:.2f}',
        'ltp': '{:.2f}',
        'StrikePrice': '{:.2f}'
    }

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Upper Strike (CE)")
        st.dataframe(
            calls_df[display_cols].style
            .map(color_change, subset=['change %'])
            .map(color_tgt_hit, subset=['TGT HIT'])
            .format(format_dict)
            .set_properties(**{'font-weight': '600', 'text-align': 'center', 'font-size': '16px'}),
            hide_index=True,
            width='stretch',
            height=1800
        )

    with col2:
        st.subheader("Lower Strike (PE)")
        st.dataframe(
            puts_df[display_cols].style
            .map(color_change, subset=['change %'])
            .map(color_tgt_hit, subset=['TGT HIT'])
            .format(format_dict)
            .set_properties(**{'font-weight': '600', 'text-align': 'center', 'font-size': '16px'}),
            hide_index=True,
            width='stretch',
            height=1800
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

    telegram_bot_token = st.secrets.get("TELEGRAM_BOT_TOKEN", "")
    telegram_chat_id = st.secrets.get("TELEGRAM_CHAT_ID", "")
    telegram_enabled = bool(telegram_bot_token and telegram_chat_id)

else:
    with st.sidebar:
        st.header("Configuration")

        saved_token = load_token()
        access_token = st.text_input("Upstox Access Token", value=saved_token, type="password")

        if access_token and access_token != saved_token:
            save_token(access_token)

        st.markdown("---")
        st.header("Telegram Alerts")

        telegram_enabled = st.checkbox(
            "Enable Trigger Alerts",
            value=st.session_state.get('telegram_enabled', False),
            key='telegram_enabled',
            help="Sends a Telegram message the moment an option's LTP crosses its Trigger price (change % >= 100)."
        )

        telegram_bot_token = st.text_input(
            "Bot Token",
            type="password",
            value=st.session_state.get('telegram_bot_token', ''),
            key='telegram_bot_token',
            help="Create a bot via @BotFather on Telegram to get this token."
        )

        telegram_chat_id = st.text_input(
            "Chat ID",
            value=st.session_state.get('telegram_chat_id', ''),
            key='telegram_chat_id',
            help="Your personal or group chat ID. Message @userinfobot to find yours."
        )

        tg_col1, tg_col2 = st.columns(2)
        test_telegram_clicked = tg_col1.button("Send Test", use_container_width=True)
        reset_alert_state_clicked = tg_col2.button("Reset Alerts", use_container_width=True)

        if reset_alert_state_clicked:
            save_trigger_alert_state(set())
            st.success("Alert state cleared — already-triggered options will alert again.")

        if test_telegram_clicked:
            success, error = send_telegram_alert(
                telegram_bot_token,
                telegram_chat_id,
                "✅ Test alert from OTM Positional Scanner — Telegram is wired up correctly."
            )
            if success:
                st.success("Test message sent — check Telegram.")
            else:
                st.error(f"Test message failed: {error}")

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

        # Monthly IV Excel Uploader (replaces the old Bhavcopy ZIP upload)
        st.subheader("Monthly IV Excel")
        up_iv = st.file_uploader(
            "Upload Monthly IV Excel",
            type=['xlsx'],
            key='iv_up',
            help="The output file from the IV Sheet Generator (e.g. 'Monthly IV 25AUG2026.xlsx')."
        )
        if up_iv is not None:
            with open(MONTHLY_IV_FILE, "wb") as f:
                f.write(up_iv.getvalue())
            save_meta('MonthlyIVFileName', up_iv.name)

            detected_expiry = extract_expiry_from_filename(up_iv.name)
            if detected_expiry is not None:
                save_meta('MonthlyIVExpiry', detected_expiry.strftime('%Y-%m-%d'))
                st.success(f"Uploaded {up_iv.name} — Expiry detected: {detected_expiry.strftime('%d-%b-%Y')}")
            else:
                st.warning(f"Uploaded {up_iv.name} — could not auto-detect expiry from filename. Please confirm it below.")

        meta = load_meta()

        if os.path.exists(MONTHLY_IV_FILE):
            st.caption(f"📄 File: {meta.get('MonthlyIVFileName', 'monthly_iv.xlsx')}")

            stored_expiry_str = meta.get('MonthlyIVExpiry')
            try:
                default_expiry_date = datetime.strptime(stored_expiry_str, '%Y-%m-%d').date() if stored_expiry_str else get_ist_now().date()
            except Exception:
                default_expiry_date = get_ist_now().date()

            manual_expiry = st.date_input(
                "Confirm Expiry Date",
                value=default_expiry_date,
                help="Must match the option expiry exactly, so it can be matched against NSE.json."
            )
            save_meta('MonthlyIVExpiry', manual_expiry.strftime('%Y-%m-%d'))

        st.markdown("---")
        st.header("Auto Refresh")
        auto_refresh = st.checkbox("Enable Auto-Refresh", value=False)
        refresh_interval = st.slider("Refresh Interval (seconds)", min_value=5, max_value=60, value=15)

# --- Main Page ---
st.title("OTM Positional Scanner")

nse_json_df = load_nse_json()

if not nse_json_df.empty:
    st.header("Monthly OTM Options")

    meta = load_meta()
    expiry_str = meta.get('MonthlyIVExpiry')
    target_expiry = None
    if expiry_str:
        try:
            target_expiry = pd.to_datetime(expiry_str).normalize()
        except Exception:
            target_expiry = None

    if os.path.exists(MONTHLY_IV_FILE) and target_expiry is not None:
        st.info(f"📅 Displaying Expiry: **{target_expiry.strftime('%d-%b-%Y')}**")

        run_every = refresh_interval if auto_refresh else None

        @st.fragment(run_every=run_every)
        def show_monthly():
            df_m = process_iv_excel(MONTHLY_IV_FILE, nse_json_df, target_expiry)
            display_option_chain(df_m, access_token, "Monthly", telegram_enabled, telegram_bot_token, telegram_chat_id)
        show_monthly()
    else:
        st.warning("Monthly IV Excel file not found. Please upload it in the sidebar.")

else:
    st.error("Critical Error: NSE.json could not be loaded.")

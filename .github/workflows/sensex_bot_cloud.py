# -*- coding: utf-8 -*-
# Last Updated: 2026-08-22
"""Standalone Single-File Cloud Autonomous SENSEX Options Bot.

This script consolidates:
1. Quantitative Primitives (DTE mapping, BSM Delta, EPM Master Grid)
2. Angel One SmartAPI Market Data & Auth Adapter
3. Active Weekly Expiry Contract Selection
4. Excel Signal & Order Tracker (signal_tracker.xlsx)
5. Automated Telegram & Mobile SMS Push Notifications
6. Autonomous Headless & Terminal Execution Engine

Run locally or on Cloud Schedulers (GitHub Actions / Railway / PythonAnywhere):
    python sensex_bot_cloud.py
"""

from __future__ import annotations

import math
import os
import re
import sys
import time
import logging
import urllib.request
import urllib.parse
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Mapping, Iterable, Sequence

import pandas as pd
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Reconfigure standard streams to UTF-8 to prevent UnicodeEncodeError (charmap) when printing emojis in Windows consoles
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr.encoding != 'utf-8':
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

IST = timezone(timedelta(hours=5, minutes=30))

def ist_converter(*args):
    return datetime.now(IST).timetuple()

# Logging setup with IST converter to ensure running updates printing time is strictly in Indian Time (IST)
import logging
logging.Formatter.converter = ist_converter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("cloud_bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger("sensex_cloud_bot")

# Enable Virtual Terminal Processing for Windows console to support ANSI escape sequences and colors
if os.name == "nt":
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        # ENABLE_PROCESSED_OUTPUT = 0x0001, ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        pass


# =========================================================================
# 1. QUANTITATIVE MATHEMATICS & DTE MAPPING
# =========================================================================

def to_ist_datetime(value: Any = None) -> datetime:
    """Convert input date/string safely into localized IST datetime."""
    if value is None:
        return datetime.now(IST)

    if isinstance(value, str):
        value = value.strip()
        formats = (
            "%Y-%m-%d", "%Y-%b-%d", "%Y-%B-%d",
            "%d-%b-%Y", "%d%b%Y", "%d%b%y", "%d-%b-%y",
            "%Y%m%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d%m%Y"
        )
        for fmt in formats:
            for variant in (value, value.upper(), value.capitalize()):
                try:
                    dt = datetime.strptime(variant, fmt)
                    return dt.replace(tzinfo=IST)
                except ValueError:
                    continue
        return datetime.now(IST)

    if isinstance(value, datetime):
        return value.replace(tzinfo=IST) if value.tzinfo is None else value.astimezone(IST)

    if isinstance(value, date):
        return datetime(value.year, value.month, value.day, tzinfo=IST)

    return datetime.now(IST)


def calculate_dte_sqrt(expiry: Any, as_of: Any = None) -> tuple[float, float]:
    """Deterministic trading-day DTE mapping:
      Monday=4.0, Tuesday=3.0, Wednesday=2.0, Thursday=1.0, Friday/Weekend=5.0
    """
    as_of_ist = to_ist_datetime(as_of)
    wd = as_of_ist.weekday()
    if wd == 0:
        dte_days = 4.0
    elif wd == 1:
        dte_days = 3.0
    elif wd == 2:
        dte_days = 2.0
    elif wd == 3:
        dte_days = 1.0
    else:
        dte_days = 5.0

    time_factor = math.sqrt(dte_days / 365.0)
    return dte_days, time_factor


def calculate_bsm_delta(spot: float, strike: float, dte_days: float, vix: float = 13.5, option_type: str = "CE") -> float:
    """Calculate Black-Scholes Delta."""
    if spot <= 0 or strike <= 0 or dte_days <= 0:
        return 0.50 if option_type == "CE" else -0.50

    t = max(dte_days, 0.001) / 365.0
    sigma = max(vix, 1.0) / 100.0
    r = 0.07  # RBI repo rate baseline

    try:
        d1 = (math.log(spot / strike) + (r + 0.5 * sigma ** 2) * t) / (sigma * math.sqrt(t))
        norm_cdf = 0.5 * (1.0 + math.erf(d1 / math.sqrt(2.0)))
        return norm_cdf if option_type.upper() == "CE" else norm_cdf - 1.0
    except Exception:
        return 0.50 if option_type.upper() == "CE" else -0.50


@dataclass(frozen=True)
class MasterGridLeg:
    option_type: str
    strike: float
    ltp: float
    delta: float
    target_epm: float
    epm_lower_range: float
    sl_auto: float
    practical_target: float
    expiry: str = ""
    trading_symbol: str = ""


@dataclass(frozen=True)
class EPMMasterGrid:
    spot: float
    vix: float
    dte: float
    dte_sqrt: float
    index_move: float
    noise_10: float
    lower_index: float
    upper_index: float
    ce_leg: MasterGridLeg
    pe_leg: MasterGridLeg
    ce_legs: list[MasterGridLeg] = None
    pe_legs: list[MasterGridLeg] = None


def calculate_master_grid_leg(
    option_type: str,
    strike: float,
    ltp: float,
    delta: float,
    index_move: float,
    time_factor: float = 1.0,
    vix: float = 13.5,
    dte: float = 1.0,
    buffer: float = 0.15,
    expiry: str = "",
    trading_symbol: str = "",
) -> MasterGridLeg:
    abs_delta = abs(float(delta))
    ltp_val = float(ltp)
    buf_val = float(buffer)

    epm_val = ltp_val * abs_delta * time_factor * (vix / 100.0)
    epm_range = ltp_val * epm_val

    epm_lower_range = ltp_val - (epm_range * abs_delta * buf_val)
    target_epm = ltp_val + (epm_range * abs_delta * buf_val)
    sl_auto = epm_lower_range - 15.0
    practical_target = ltp_val + (index_move * abs_delta * 0.21)

    return MasterGridLeg(
        option_type=str(option_type).upper(),
        strike=float(strike),
        ltp=ltp_val,
        delta=abs_delta,
        target_epm=target_epm,
        epm_lower_range=epm_lower_range,
        sl_auto=sl_auto,
        practical_target=practical_target,
        expiry=str(expiry),
        trading_symbol=str(trading_symbol),
    )


def calculate_master_grid(
    spot: float, vix: float, dte: float,
    ce_ltp: float, ce_delta: float, ce_strike: float,
    pe_ltp: float, pe_delta: float, pe_strike: float,
    buffer: float = 0.12,
    ce_legs_data: list[Any] | None = None,
    pe_legs_data: list[Any] | None = None,
    ce_expiry: str = "", ce_trading_symbol: str = "",
    pe_expiry: str = "", pe_trading_symbol: str = "",
) -> EPMMasterGrid:
    spot_val = float(spot)
    vix_val = float(vix) if float(vix) > 0 else 13.5
    dte_val = float(dte)
    time_factor = math.sqrt(max(0.0, dte_val) / 365.0)

    index_move = spot_val * (vix_val / 100.0) * time_factor
    noise_10 = index_move * 0.10
    lower_index = spot_val - index_move
    upper_index = spot_val + index_move

    ce_leg = calculate_master_grid_leg("CE", ce_strike, ce_ltp, ce_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=ce_expiry, trading_symbol=ce_trading_symbol)
    pe_leg = calculate_master_grid_leg("PE", pe_strike, pe_ltp, pe_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=pe_expiry, trading_symbol=pe_trading_symbol)

    ce_legs = []
    if ce_legs_data:
        for item in ce_legs_data:
            c_ltp, c_delta, c_strike = item[0], item[1], item[2]
            c_exp = item[3] if len(item) > 3 else ce_expiry
            c_sym = item[4] if len(item) > 4 else ce_trading_symbol
            ce_legs.append(calculate_master_grid_leg("CE", c_strike, c_ltp, c_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=c_exp, trading_symbol=c_sym))
    else:
        ce_legs = [ce_leg]

    pe_legs = []
    if pe_legs_data:
        for item in pe_legs_data:
            p_ltp, p_delta, p_strike = item[0], item[1], item[2]
            p_exp = item[3] if len(item) > 3 else pe_expiry
            p_sym = item[4] if len(item) > 4 else pe_trading_symbol
            pe_legs.append(calculate_master_grid_leg("PE", p_strike, p_ltp, p_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=p_exp, trading_symbol=p_sym))
    else:
        pe_legs = [pe_leg]

    return EPMMasterGrid(
        spot=spot_val, vix=vix_val, dte=dte_val, dte_sqrt=time_factor,
        index_move=index_move, noise_10=noise_10,
        lower_index=lower_index, upper_index=upper_index,
        ce_leg=ce_leg, pe_leg=pe_leg,
        ce_legs=ce_legs, pe_legs=pe_legs,
    )


# =========================================================================
# 2. MOBILE NOTIFICATIONS (TELEGRAM / TELEGRAM BOT / SMS WEBHOOK)
# =========================================================================

class TelegramCommandListener:
    """Robust Telegram command listener that flushes old chat history on startup."""

    def __init__(self, bot_token: str | None) -> None:
        self.bot_token = bot_token
        self.last_update_id = 0
        self.startup_time = int(time.time())
        self._flush_old_updates()

    def _flush_old_updates(self) -> None:
        """Flush old chat history on startup so past STOP messages never trigger false stops."""
        if not self.bot_token:
            return
        try:
            import json
            url = f"https://api.telegram.org/bot{self.bot_token}/getUpdates?offset=-1"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode())
                if data.get("ok") and data.get("result"):
                    self.last_update_id = data["result"][-1].get("update_id", 0)
                    flush_url = f"https://api.telegram.org/bot{self.bot_token}/getUpdates?offset={self.last_update_id + 1}"
                    urllib.request.urlopen(urllib.request.Request(flush_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=3)
                    logger.info("📱 Telegram updates flushed on startup (Last Update ID: %d)", self.last_update_id)
        except Exception as e:
            logger.debug("Failed to flush Telegram updates: %s", e)

    def get_new_command(self) -> tuple[str | None, int | None]:
        """Poll ONLY for NEW incoming messages arriving AFTER bot startup."""
        if not self.bot_token:
            return None, None
        try:
            import json
            url = f"https://api.telegram.org/bot{self.bot_token}/getUpdates?offset={self.last_update_id + 1}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode())
                if data.get("ok") and data.get("result"):
                    for update in data["result"]:
                        u_id = update.get("update_id", 0)
                        if u_id > self.last_update_id:
                            self.last_update_id = u_id
                            msg_obj = update.get("message") or update.get("channel_post") or {}
                            
                            # Ignore past messages sent before bot startup to prevent false triggers
                            msg_date = msg_obj.get("date", 0)
                            if msg_date < self.startup_time:
                                continue

                            msg_text = str(msg_obj.get("text") or "").strip().upper()
                            parts = msg_text.split()
                            first_word = parts[0] if parts else ""

                            if first_word in ("STOP", "HALT", "EXIT", "CLOSE", "/STOP"):
                                return "STOP", None
                            elif first_word in ("LIVE", "REAL", "/LIVE"):
                                lots = 1
                                if len(parts) > 1 and parts[1].isdigit():
                                    lots = max(1, int(parts[1]))
                                return "LIVE", lots
                            elif first_word in ("DEMO", "PAPER", "/DEMO"):
                                return "DEMO", None
                            elif first_word == "ADD":
                                lots = 1
                                if len(parts) > 1 and parts[1].isdigit():
                                    lots = max(1, int(parts[1]))
                                return "ADD", lots
                            elif first_word == "BUFFER":
                                points = 5.0
                                if len(parts) > 1:
                                    try:
                                        points = max(1.0, float(parts[1]))
                                    except ValueError:
                                        pass
                                return "BUFFER", points
                            elif first_word == "SL":
                                price = 0.0
                                if len(parts) > 1:
                                    try:
                                        price = max(1.0, float(parts[1]))
                                    except ValueError:
                                        pass
                                return "SL", price
        except Exception:
            pass
        return None, None


def submit_angel_order(smart_api: Any, trading_symbol: str, symbol_token: str, transaction_type: str = "BUY", quantity: int = 10) -> Any:
    """Submit real Market Order to Angel One SmartAPI."""
    try:
        order_params = {
            "variety": "NORMAL",
            "tradingsymbol": trading_symbol,
            "symboltoken": symbol_token,
            "transactiontype": transaction_type,
            "exchange": "BFO",
            "ordertype": "MARKET",
            "producttype": "INTRADAY",
            "duration": "DAY",
            "price": "0",
            "squareoff": "0",
            "stoploss": "0",
            "quantity": str(quantity),
        }
        order_id = smart_api.placeOrder(order_params)
        logger.info("⚡ [REAL ORDER SUBMITTED] %s %d %s | Order ID: %s", transaction_type, quantity, trading_symbol, order_id)
        send_mobile_alert(f"🚨 *REAL ORDER PLACED ON ANGEL ONE*\n\nAction: *{transaction_type}*\nContract: *{trading_symbol}*\nQuantity: *{quantity}*\nOrder ID: `{order_id}`")
        return order_id
    except Exception as exc:
        logger.error("❌ Real Order Submission Failed: %s", exc)
        send_mobile_alert(f"⚠️ *ORDER SUBMISSION ERROR*\nFailed to place {transaction_type} for {trading_symbol}: {exc}")
        return None


def send_mobile_alert(message: str) -> None:
    """Send mobile push notifications via Telegram Bot API or CallMeBot WhatsApp API."""
    # 1. Telegram Push Notification
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if bot_token and chat_id:
        try:
            url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
            payload = urllib.parse.urlencode({"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}).encode("utf-8")
            req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"})
            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status == 200:
                    logger.info("📱 Telegram push notification sent successfully.")
        except Exception as e:
            logger.warning("Failed to send Telegram alert: %s", e)

    # 2. WhatsApp Notification via CallMeBot API
    wa_phone = os.environ.get("WHATSAPP_PHONE")
    wa_apikey = os.environ.get("WHATSAPP_API_KEY")

    if wa_phone and wa_apikey:
        try:
            clean_msg = message.replace("*", "").replace("`", "")
            encoded_text = urllib.parse.quote(clean_msg)
            url = f"https://api.callmebot.com/whatsapp.php?phone={wa_phone}&text={encoded_text}&apikey={wa_apikey}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=8) as response:
                if response.status == 200:
                    logger.info("🟢 WhatsApp notification sent successfully.")
        except Exception as e:
            logger.warning("Failed to send WhatsApp alert: %s", e)


# =========================================================================
# 3. EXCEL TRACKER LOGGING
# =========================================================================

class ExcelTracker:
    def __init__(self, filename: str = "signal_tracker.xlsx") -> None:
        self.filename = filename
        self.signals_list: list[dict[str, Any]] = []
        self.orders_list: list[dict[str, Any]] = []
        self._load_existing()

    def _load_existing(self) -> None:
        if os.path.exists(self.filename):
            try:
                with pd.ExcelFile(self.filename) as xls:
                    if "Signal Tracker" in xls.sheet_names:
                        self.signals_list = pd.read_excel(xls, "Signal Tracker").to_dict("records")
                    if "Demo Order Tracker" in xls.sheet_names:
                        self.orders_list = pd.read_excel(xls, "Demo Order Tracker").to_dict("records")
            except Exception as e:
                logger.warning("Could not read existing Excel tracker file: %s", e)

    def add_signal(self, row: dict[str, Any]) -> None:
        self.signals_list.append(row)
        self._save()

    def add_order(self, row: dict[str, Any]) -> None:
        self.orders_list.append(row)
        self._save()

    def _save(self) -> None:
        try:
            df_sig = pd.DataFrame(self.signals_list)
            df_ord = pd.DataFrame(self.orders_list)
            with pd.ExcelWriter(self.filename, engine="openpyxl") as writer:
                df_sig.to_excel(writer, sheet_name="Signal Tracker", index=False)
                df_ord.to_excel(writer, sheet_name="Demo Order Tracker", index=False)
        except Exception as e:
            logger.warning("Could not save Excel tracker: %s", e)


# =========================================================================
# 4. CONTRACT SELECTION & SMARTAPI ADAPTER
# =========================================================================

@dataclass(frozen=True)
class OptionContract:
    exchange: str
    trading_symbol: str
    symbol_token: str
    expiry: Any
    strike: float
    option_type: str
    delta: float


def select_itm_contracts(
    contracts: Iterable[OptionContract],
    spot_price: float,
    option_type: str,
    count: int = 3,
) -> list[OptionContract]:
    spot = float(spot_price)
    matching = [c for c in contracts if c.option_type == option_type]
    if not matching:
        raise ValueError(f"No matching contracts found for {option_type}")

    now_ist = datetime.now(IST)
    contracts_with_expiry = []
    for c in matching:
        try:
            exp_dt = to_ist_datetime(c.expiry)
            if exp_dt.date() >= now_ist.date():
                contracts_with_expiry.append((c, exp_dt.date()))
        except Exception:
            pass

    if contracts_with_expiry:
        earliest_expiry = min(exp_dt_date for _, exp_dt_date in contracts_with_expiry)
        matching = [c for c, exp_dt_date in contracts_with_expiry if exp_dt_date == earliest_expiry]

    # Deduplicate contracts by unique strike so we pick 3 distinct strike levels
    by_strike: dict[float, OptionContract] = {}
    for c in matching:
        if c.strike not in by_strike:
            by_strike[c.strike] = c

    unique_contracts = list(by_strike.values())

    if option_type == "CE":
        itm = sorted([c for c in unique_contracts if c.strike <= spot], key=lambda c: c.strike, reverse=True)
        if len(itm) < count:
            all_sorted = sorted(unique_contracts, key=lambda c: c.strike, reverse=True)
            return all_sorted[:count]
        return itm[:count]
    else:
        itm = sorted([c for c in unique_contracts if c.strike >= spot], key=lambda c: c.strike)
        if len(itm) < count:
            all_sorted = sorted(unique_contracts, key=lambda c: c.strike)
            return all_sorted[:count]
        return itm[:count]


def select_nearest_itm_contract(
    contracts: Iterable[OptionContract],
    spot_price: float,
    option_type: str,
) -> OptionContract:
    res = select_itm_contracts(contracts, spot_price, option_type, count=1)
    return res[0]


def create_authenticated_smartapi_client() -> Any:
    req_keys = ("ANGEL_ONE_API_KEY", "ANGEL_ONE_CLIENT_CODE", "ANGEL_ONE_PASSWORD", "ANGEL_ONE_TOTP_SECRET")
    missing = [k for k in req_keys if not os.environ.get(k)]
    if missing:
        raise RuntimeError("Missing Angel One secret(s): " + ", ".join(missing))

    import pyotp
    from SmartApi import SmartConnect

    smart_api = SmartConnect(api_key=os.environ["ANGEL_ONE_API_KEY"])
    login_response = smart_api.generateSession(
        os.environ["ANGEL_ONE_CLIENT_CODE"],
        os.environ["ANGEL_ONE_PASSWORD"],
        pyotp.TOTP(os.environ["ANGEL_ONE_TOTP_SECRET"]).now(),
    )
    if not isinstance(login_response, dict) or login_response.get("status") is not True:
        raise RuntimeError(f"Angel One authentication failed: {login_response}")
        
    # Attach session keys dynamically for WebSocket use
    try:
        smart_api.feed_token = login_response["data"]["feedToken"]
        smart_api.auth_token = login_response["data"]["jwtToken"]
    except Exception:
        smart_api.feed_token = None
        smart_api.auth_token = None
        
    return smart_api


# =========================================================================
# 5. MAIN CLOUD RUNNER
# =========================================================================

import threading

class LiveWSFeed:
    def __init__(self, client_code: str, feed_token: str, api_key: str, auth_token: str):
        self.client_code = client_code
        self.feed_token = feed_token
        self.api_key = api_key
        self.auth_token = auth_token
        self.prices: dict[str, float] = {}
        self.ws = None
        self.thread = None
        self.is_connected = False

    def on_data(self, ws, message):
        try:
            if isinstance(message, list):
                for tick in message:
                    token = str(tick.get("token") or "")
                    last_traded_price = tick.get("last_traded_price") or tick.get("ltp")
                    if token and last_traded_price is not None:
                        # Price is in paisa if > 1000000, else standard float
                        val = float(last_traded_price)
                        self.prices[token] = val / 100.0 if val > 1000000 else val
            elif isinstance(message, dict):
                token = str(message.get("token") or "")
                last_traded_price = message.get("last_traded_price") or message.get("ltp")
                if token and last_traded_price is not None:
                    val = float(last_traded_price)
                    self.prices[token] = val / 100.0 if val > 1000000 else val
        except Exception:
            pass

    def on_open(self, ws):
        self.is_connected = True
        logger.info("🟢 WebSocket Connection Opened successfully!")

    def on_close(self, ws, close_status_code, close_msg):
        self.is_connected = False
        logger.info("🔴 WebSocket Connection Closed.")

    def on_error(self, ws, error):
        self.is_connected = False
        logger.debug("WebSocket Error: %s", error)

    def start(self, tokens_to_subscribe: list[str]):
        """Start the WebSocket in a background thread to update prices continuously."""
        try:
            from SmartApi.smartWebSocketV2 import SmartWebSocketV2
            self.ws = SmartWebSocketV2(self.auth_token, self.api_key, self.client_code, self.feed_token)
            
            self.ws.on_open = self.on_open
            self.ws.on_data = self.on_data
            self.ws.on_error = self.on_error
            self.ws.on_close = self.on_close
            
            correlation_id = "sensex_bot_feed"
            action = 1  # Subscribe
            mode = 3    # Full/LTP mode
            
            subscription_list = []
            for token in tokens_to_subscribe:
                # 1 = NSECM, 2 = NSEFO, 3 = BSECM, 4 = BFO
                exchange = 4 if len(token) > 6 else 3
                subscription_list.append({
                    "exchangeType": exchange,
                    "tokens": [token]
                })

            def run_ws():
                try:
                    self.ws.connect()
                    self.ws.subscribe(correlation_id, mode, subscription_list)
                except Exception as e:
                    logger.debug("WebSocket run exception: %s", e)
                    self.is_connected = False

            self.thread = threading.Thread(target=run_ws, daemon=True)
            self.thread.start()
        except Exception as e:
            logger.debug("Could not initialize SmartWebSocketV2: %s", e)
            self.is_connected = False


def get_current_15m_candle_ohl(smart_api: Any, exchange: str, symbol_token: str) -> tuple[float | None, float | None]:
    """Fetch the current 15-minute candle's Open and Low prices from SmartAPI.
    Returns (open, low) on success, or (None, None) on error.
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(minutes=45)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = smart_api.getCandle(params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if candles:
                last_candle = candles[-1]
                if len(last_candle) >= 4:
                    c_open = float(last_candle[1])
                    c_low = float(last_candle[3])
                    return c_open, c_low
    except Exception as e:
        logger.debug("Error fetching 15m candle OHLC: %s", e)
    return None, None


def get_15m_mfi(smart_api: Any, exchange: str, symbol_token: str, period: int = 5) -> tuple[float, float, float | None]:
    """Calculate Money Flow Index (MFI) on 15-minute timeframe.
    Returns (curr_mfi, prev_mfi, prev_candle_low).
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(minutes=15 * (period + 10))
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = smart_api.getCandle(params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= period + 1:
                typical_prices = []
                volumes = []
                lows = []
                for c in candles:
                    if len(c) >= 6:
                        h, l, cl, v = float(c[2]), float(c[3]), float(c[4]), float(c[5])
                        typical_prices.append((h + l + cl) / 3.0)
                        volumes.append(v if v > 0 else 1.0)
                        lows.append(l)

                if len(typical_prices) >= period + 1:
                    def calc_mfi_at(end_idx):
                        pos_mf = 0.0
                        neg_mf = 0.0
                        for i in range(end_idx - period + 1, end_idx + 1):
                            raw_mf = typical_prices[i] * volumes[i]
                            if typical_prices[i] > typical_prices[i - 1]:
                                pos_mf += raw_mf
                            elif typical_prices[i] < typical_prices[i - 1]:
                                neg_mf += raw_mf
                        if pos_mf == 0.0:
                            return 0.0
                        if neg_mf == 0.0:
                            return 100.0
                        mfr = pos_mf / neg_mf
                        return 100.0 - (100.0 / (1.0 + mfr))

                    curr_mfi = calc_mfi_at(len(typical_prices) - 1)
                    prev_mfi = calc_mfi_at(len(typical_prices) - 2)
                    prev_low = lows[-2] if len(lows) >= 2 else lows[-1]
                    return curr_mfi, prev_mfi, prev_low
    except Exception as e:
        logger.debug("Error calculating 15m MFI: %s", e)
    return 50.0, 50.0, None


def get_current_time_slot() -> str:
    import json
    now_dt = datetime.now(IST)
    m_of_day = now_dt.hour * 60 + now_dt.minute
    if m_of_day < (9 * 60 + 45):  # Before 9:45 AM
        return "09:15"
    elif m_of_day < (12 * 60 + 15):  # Before 12:15 PM
        return "09:45"
    else:
        return "12:15"


def load_bot_memory() -> dict | None:
    import json
    path = Path("bot_state_memory.json")
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        today_str = datetime.now(IST).strftime("%Y-%m-%d")
        if data.get("date") == today_str:
            return data
    except Exception as e:
        logger.warning("Failed to load bot memory: %s", e)
    return None


def save_bot_memory_full(trades_completed: int, grid_time_slot: str, grid: EPMMasterGrid, ce_contract: OptionContract, pe_contract: OptionContract, bot_state: str, active_contract: OptionContract | None, active_entry_price: float, active_sl: float, active_target: float, peak_price: float, trailing_active: bool, offloaded: bool, lot_size: int, entry_time: datetime | None) -> None:
    import json
    try:
        data = {
            "date": datetime.now(IST).strftime("%Y-%m-%d"),
            "trades_completed": trades_completed,
            "grid_time_slot": grid_time_slot,
            "bot_state": bot_state,
            "active_entry_price": active_entry_price,
            "active_sl": active_sl,
            "active_target": active_target,
            "peak_price": peak_price,
            "trailing_active": trailing_active,
            "offloaded": offloaded,
            "lot_size": lot_size,
            "entry_time": entry_time.isoformat() if entry_time else None,
            "grid": {
                "spot": grid.spot,
                "vix": grid.vix,
                "dte": grid.dte,
                "index_move": grid.index_move,
                "ce_leg": {
                    "strike": grid.ce_leg.strike,
                    "ltp": grid.ce_leg.ltp,
                    "delta": grid.ce_leg.delta,
                    "target_epm": grid.ce_leg.target_epm,
                    "epm_lower_range": grid.ce_leg.epm_lower_range,
                    "sl_auto": grid.ce_leg.sl_auto,
                    "practical_target": grid.ce_leg.practical_target,
                    "option_type": grid.ce_leg.option_type
                },
                "pe_leg": {
                    "strike": grid.pe_leg.strike,
                    "ltp": grid.pe_leg.ltp,
                    "delta": grid.pe_leg.delta,
                    "target_epm": grid.pe_leg.target_epm,
                    "epm_lower_range": grid.pe_leg.epm_lower_range,
                    "sl_auto": grid.pe_leg.sl_auto,
                    "practical_target": grid.pe_leg.practical_target,
                    "option_type": grid.pe_leg.option_type
                },
                "ce_legs": [
                    {
                        "strike": leg.strike,
                        "ltp": leg.ltp,
                        "delta": leg.delta,
                        "target_epm": leg.target_epm,
                        "epm_lower_range": leg.epm_lower_range,
                        "sl_auto": leg.sl_auto,
                        "practical_target": leg.practical_target,
                        "option_type": leg.option_type
                    } for leg in (grid.ce_legs or [grid.ce_leg])
                ],
                "pe_legs": [
                    {
                        "strike": leg.strike,
                        "ltp": leg.ltp,
                        "delta": leg.delta,
                        "target_epm": leg.target_epm,
                        "epm_lower_range": leg.epm_lower_range,
                        "sl_auto": leg.sl_auto,
                        "practical_target": leg.practical_target,
                        "option_type": leg.option_type
                    } for leg in (grid.pe_legs or [grid.pe_leg])
                ]
            },
            "ce_contract": {
                "exchange": ce_contract.exchange,
                "trading_symbol": ce_contract.trading_symbol,
                "symbol_token": ce_contract.symbol_token,
                "expiry": ce_contract.expiry,
                "strike": ce_contract.strike,
                "option_type": ce_contract.option_type,
                "delta": ce_contract.delta
            },
            "pe_contract": {
                "exchange": pe_contract.exchange,
                "trading_symbol": pe_contract.trading_symbol,
                "symbol_token": pe_contract.symbol_token,
                "expiry": pe_contract.expiry,
                "strike": pe_contract.strike,
                "option_type": pe_contract.option_type,
                "delta": pe_contract.delta
            }
        }
        with open("bot_state_memory.json", "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        logger.info("💾 Bot state and active trade memory saved successfully.")
    except Exception as e:
        logger.warning("Failed to save bot active memory: %s", e)


def save_bot_memory(trades_completed: int, grid_time_slot: str, grid: EPMMasterGrid, ce_contract: OptionContract, pe_contract: OptionContract) -> None:
    save_bot_memory_full(trades_completed, grid_time_slot, grid, ce_contract, pe_contract, "IDLE", None, 0.0, 0.0, 0.0, 0.0, False, False, 1, None)


def get_current_1m_candles(smart_api: Any, exchange: str, symbol_token: str, count_mins: int = 15) -> list:
    """Fetch recent 1-minute candles from SmartAPI.
    Each candle is: [timestamp, open, high, low, close, volume]
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(minutes=count_mins)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "ONE_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = smart_api.getCandle(params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            return res["data"]
    except Exception as e:
        logger.debug("Error fetching 1m candles: %s", e)
    return []


def load_delta_map() -> dict[str, dict[str, Any]]:
    """Load official Angel One Script Master metadata map from local files if available."""
    for path in ("delta_map.json", "../../delta_map.json", "0_sensex_options_delta_1786941098090.json", "../../0_sensex_options_delta_1786941098090.json"):
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
                    elif isinstance(data, list):
                        res = {}
                        for row in data:
                            sym = row.get("trading_symbol") or row.get("tradingsymbol")
                            if sym:
                                res[sym] = row
                        return res
            except Exception:
                pass
    return {}


def build_epm_grid_and_contracts(
    smart_api: Any,
    current_slot: str,
    spot_price: float,
    spot_open: float,
    vix_val: float,
) -> tuple[EPMMasterGrid, list[OptionContract], list[OptionContract]]:
    """Fetch option contracts, select top 3 ITM strikes for CE and PE, calculate EPM grid for all of them."""
    search_res = smart_api.searchScrip("BFO", "SENSEX")
    rows = search_res.get("data", []) if isinstance(search_res, dict) else []

    delta_map = load_delta_map()
    contracts: list[OptionContract] = []
    month_map = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9, "O": 10, "N": 11, "D": 12, "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}

    for r in rows:
        symbol = str(r.get("tradingsymbol") or "").strip()
        if not symbol or not (symbol.endswith("CE") or symbol.endswith("PE")):
            continue
        opt_type = "CE" if symbol.endswith("CE") else "PE"
        token = str(r.get("symboltoken") or "").strip()

        # 1. Prefer API provided expiry or Script Master metadata expiry directly
        raw_exp = str(r.get("expiry") or "").strip()
        if not raw_exp and symbol in delta_map:
            raw_exp = str(delta_map[symbol].get("expiry") or "").strip()

        expiry_val = None
        if raw_exp:
            try:
                parsed = to_ist_datetime(raw_exp)
                expiry_val = parsed.strftime("%Y-%m-%d")
            except Exception:
                expiry_val = raw_exp

        strike_val = None
        if symbol in delta_map and "strike" in delta_map[symbol]:
            try:
                strike_val = float(delta_map[symbol]["strike"])
            except Exception:
                pass

        m_sym = re.search(r"(?:BSE)?SENSEX(\d{2})([A-Za-z]{3}|\d|[ONDond])(?:(0[1-9]|[12][0-9]|3[01]))?(\d{4,6})(CE|PE)$", symbol, re.IGNORECASE)
        if m_sym:
            yy, m_str, dd, str_val, _ = m_sym.groups()
            if not strike_val:
                strike_val = float(str_val)
            if not expiry_val:
                m_num = month_map.get(m_str.upper(), 8)
                if dd:
                    expiry_val = f"20{yy}-{m_num:02d}-{int(dd):02d}"
                else:
                    expiry_val = f"20{yy}-{m_num:02d}-28"
        else:
            m_strike = re.search(r"(\d+)(?:CE|PE)$", symbol)
            if not m_strike:
                continue
            num_str = m_strike.group(1)
            if not strike_val:
                strike_val = float(num_str[-5:]) if len(num_str) >= 5 else float(num_str)

            if not expiry_val:
                m_ymd = re.search(r"(\d{4})(\d{2})(\d{2})", symbol)
                if m_ymd:
                    expiry_val = f"{m_ymd.group(1)}-{m_ymd.group(2)}-{m_ymd.group(3)}"

        if not expiry_val or not strike_val:
            continue

        try:
            exp_dt = to_ist_datetime(expiry_val)
            if exp_dt.date() < datetime.now(IST).date():
                continue
        except Exception:
            continue
            continue

        logger.debug("Parsed expiry for %s -> %s (raw: %s)", symbol, expiry_val, raw_exp)

        dte_days, _ = calculate_dte_sqrt(expiry_val)
        delta_val = calculate_bsm_delta(spot_price, strike_val, dte_days, vix_val, opt_type)

        try:
            c_obj = OptionContract("BFO", symbol, token, expiry_val, strike_val, opt_type, delta_val)
            contracts.append(c_obj)
        except Exception:
            continue

    ce_contracts = select_itm_contracts(contracts, spot_price, "CE", count=3)
    pe_contracts = select_itm_contracts(contracts, spot_price, "PE", count=3)

    ce_legs_data = []
    for c in ce_contracts:
        res = smart_api.ltpData("BFO", c.trading_symbol, c.symbol_token)
        ltp = float(res["data"]["ltp"]) if isinstance(res, dict) and res.get("data") else 500.0
        c_open = float(res["data"]["open"]) if isinstance(res, dict) and res.get("data") and res["data"].get("open") else ltp
        price_to_use = c_open if current_slot == "09:15" else ltp
        ce_legs_data.append((price_to_use, abs(c.delta), c.strike, str(c.expiry), c.trading_symbol))

    pe_legs_data = []
    for p in pe_contracts:
        res = smart_api.ltpData("BFO", p.trading_symbol, p.symbol_token)
        ltp = float(res["data"]["ltp"]) if isinstance(res, dict) and res.get("data") else 300.0
        p_open = float(res["data"]["open"]) if isinstance(res, dict) and res.get("data") and res["data"].get("open") else ltp
        price_to_use = p_open if current_slot == "09:15" else ltp
        pe_legs_data.append((price_to_use, abs(p.delta), p.strike, str(p.expiry), p.trading_symbol))

    dte_days, _ = calculate_dte_sqrt(ce_contracts[0].expiry)
    spot_to_use = spot_open if current_slot == "09:15" else spot_price

    grid = calculate_master_grid(
        spot=spot_to_use, vix=vix_val, dte=dte_days,
        ce_ltp=ce_legs_data[0][0], ce_delta=ce_legs_data[0][1], ce_strike=ce_legs_data[0][2],
        pe_ltp=pe_legs_data[0][0], pe_delta=pe_legs_data[0][1], pe_strike=pe_legs_data[0][2],
        buffer=0.10,
        ce_legs_data=ce_legs_data,
        pe_legs_data=pe_legs_data,
        ce_expiry=str(ce_contracts[0].expiry),
        ce_trading_symbol=ce_contracts[0].trading_symbol,
        pe_expiry=str(pe_contracts[0].expiry),
        pe_trading_symbol=pe_contracts[0].trading_symbol,
    )

    return grid, ce_contracts, pe_contracts


def format_grid_notification(grid: EPMMasterGrid, title: str, spot_price: float, spot_open: float, vix_val: float, dte_days: float, current_slot: str = "09:15") -> str:
    """Format clean Telegram and log notification containing all 2-3 ITM strikes."""
    price_label = "9:15 AM Open Price" if current_slot == "09:15" else f"{current_slot} Slot Price"
    lines = [
        f"{title} (Slot: {current_slot} IST)\n",
        f"📈 *Spot LTP*: ₹{spot_price:.2f} (Open: ₹{spot_open:.2f}) | *VIX*: {vix_val:.2f}%",
        f"📅 *DTE*: {dte_days:.1f} Days | *Expected Move*: ±₹{grid.index_move:.2f}\n",
        "🟢 *CE ITM STRIKES (STATIC EPM)*:"
    ]
    for idx, leg in enumerate(grid.ce_legs or [grid.ce_leg], 1):
        exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
        lines.append(
            f"• *CE {int(leg.strike)} (ITM {idx}{exp_info})*: {price_label} ₹{leg.ltp:.2f} | Delta {leg.delta:.3f}\n"
            f"  └ EPM Low: ₹{leg.epm_lower_range:.2f} | Upper Target: ₹{leg.target_epm:.2f} | Auto SL: ₹{leg.sl_auto:.2f}"
        )

    lines.append("\n🔴 *PE ITM STRIKES (STATIC EPM)*:")
    for idx, leg in enumerate(grid.pe_legs or [grid.pe_leg], 1):
        exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
        lines.append(
            f"• *PE {int(leg.strike)} (ITM {idx}{exp_info})*: {price_label} ₹{leg.ltp:.2f} | Delta {-leg.delta:.3f}\n"
            f"  └ EPM Low: ₹{leg.epm_lower_range:.2f} | Upper Target: ₹{leg.target_epm:.2f} | Auto SL: ₹{leg.sl_auto:.2f}"
        )

    return "\n".join(lines)


def get_current_5m_candles(smart_api: Any, exchange: str, symbol_token: str) -> list:
    """Fetch recent 5-minute candles from SmartAPI.
    Each candle is: [timestamp, open, high, low, close, volume]
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(minutes=30)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIVE_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = smart_api.getCandle(params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            return res["data"]
    except Exception as e:
        logger.debug("Error fetching 5m candles: %s", e)
    return []


def check_active_position_qty(smart_api: Any, symbol_token: str) -> int | None:
    """Fetch active positions from Angel One SmartAPI and return the net quantity for the given token.
    Returns None if the API call fails, to prevent false positive manual closure triggers.
    """
    try:
        res = smart_api.position()
        if isinstance(res, dict) and res.get("status") is True:
            positions_list = res.get("data")
            if positions_list is not None:
                for pos in positions_list:
                    token = str(pos.get("symboltoken") or pos.get("token") or "").strip()
                    if token == str(symbol_token).strip():
                        return abs(int(pos.get("netqty") or 0))
                return 0  # Contract not found in position list means quantity is 0
    except Exception as e:
        logger.warning("Error fetching active positions from SmartAPI: %s", e)
    return None


def execute_failsafe_sell(smart_api: Any, trading_symbol: str, symbol_token: str, quantity: int, ltp: float) -> Any:
    """Submit a MARKET sell order first. If it fails, immediately place a LIMIT sell order
    at a lower price (LTP - 10 points) to guarantee immediate execution as a marketable limit order.
    """
    # 1. Try Market Order
    try:
        order_params = {
            "variety": "NORMAL",
            "tradingsymbol": trading_symbol,
            "symboltoken": symbol_token,
            "transactiontype": "SELL",
            "exchange": "BFO",
            "ordertype": "MARKET",
            "producttype": "INTRADAY",
            "duration": "DAY",
            "price": "0",
            "squareoff": "0",
            "stoploss": "0",
            "quantity": str(quantity),
        }
        order_id = smart_api.placeOrder(order_params)
        if order_id:
            logger.info("⚡ [MARKET SELL ORDER SUCCESS] Order ID: %s", order_id)
            return order_id
    except Exception as exc:
        logger.warning("Market sell failed, attempting failsafe Limit sell: %s", exc)
    
    # 2. Try Failsafe Limit Order (Sell at LTP - 10 points to guarantee execution)
    try:
        limit_price = max(2.0, float(ltp) - 10.0)
        limit_price_str = f"{limit_price:.2f}"
        
        order_params = {
            "variety": "NORMAL",
            "tradingsymbol": trading_symbol,
            "symboltoken": symbol_token,
            "transactiontype": "SELL",
            "exchange": "BFO",
            "ordertype": "LIMIT",
            "producttype": "INTRADAY",
            "duration": "DAY",
            "price": limit_price_str,
            "squareoff": "0",
            "stoploss": "0",
            "quantity": str(quantity),
        }
        order_id = smart_api.placeOrder(order_params)
        logger.info("⚡ [FAILSAFE LIMIT SELL ORDER PLACED] Price: %s | Order ID: %s", limit_price_str, order_id)
        return order_id
    except Exception as exc:
        logger.error("❌ Failsafe Sell Failed: %s", exc)
        send_mobile_alert(f"⚠️ *CRITICAL: SELL ORDER FAILED*\nCould not execute sell for {trading_symbol}. Please close manually!")
        return None


def run_cloud_bot() -> None:
    logger.info("🚀 Starting Standalone Cloud SENSEX Options Bot...")
    excel_tracker = ExcelTracker()

    smart_api = create_authenticated_smartapi_client()

    current_slot = get_current_time_slot()
    mem = load_bot_memory()

    trades_completed = 0
    grid = None
    ce_contract = None
    pe_contract = None
    spot_price = 77500.0
    spot_open = 77500.0
    vix_val = 13.5
    dte_days = 4.0
    ce_ltp = 500.0
    pe_ltp = 300.0
    grid_from_memory = False

    # Get Spot Price
    try:
        spot_res = smart_api.ltpData("BSE", "SENSEX", "99919000")
        spot_price = float(spot_res["data"]["ltp"]) if isinstance(spot_res, dict) and spot_res.get("data") else 77500.0
        spot_open = float(spot_res["data"]["open"]) if isinstance(spot_res, dict) and spot_res.get("data") and spot_res["data"].get("open") else spot_price
    except Exception:
        pass

    # Get VIX
    try:
        vix_res = smart_api.ltpData("NSE", "INDIA VIX", "99926017")
        vix_val = float(vix_res["data"]["ltp"]) if isinstance(vix_res, dict) and vix_res.get("data") else 13.5
    except Exception:
        pass

    if mem is not None:
        trades_completed = mem.get("trades_completed", 0)
        # If the saved memory belongs to the current time slot, reuse its EPM and contract details!
        if mem.get("grid_time_slot") == current_slot:
            try:
                # Reconstruct contracts
                ce_c_data = mem["ce_contract"]
                pe_c_data = mem["pe_contract"]

                ce_exp_dt = to_ist_datetime(ce_c_data.get("expiry"))
                pe_exp_dt = to_ist_datetime(pe_c_data.get("expiry"))
                now_dt = datetime.now(IST)

                ce_strike = float(ce_c_data.get("strike", 0))
                pe_strike = float(pe_c_data.get("strike", 0))

                # Validate recalled strikes: Must be valid SENSEX index strike range (> 50000)
                if ce_strike < 50000 or pe_strike < 50000:
                    raise ValueError(f"Recalled contract strike (CE: {ce_strike}, PE: {pe_strike}) is corrupted.")

                # Validate recalled expiry: Must be active future expiry (not today/past unless official expiry day)
                if ce_exp_dt.date() < now_dt.date() or (ce_exp_dt.date() == now_dt.date() and ce_exp_dt.strftime("%d%b%Y").upper() not in ("03SEP2026", "10SEP2026", "17SEP2026", "24SEP2026")):
                    raise ValueError(f"Recalled CE expiry ({ce_exp_dt.date()}) is not active future weekly expiry.")

                if (ce_exp_dt.date() - now_dt.date()).days > 10:
                    raise ValueError(f"Recalled CE expiry ({ce_exp_dt.date()}) is not current active weekly expiry.")

                logger.info("🔮 [RECALL MEMORY] Recalled last stored EPM for slot %s (Expiry: %s).", current_slot, ce_exp_dt.strftime("%Y-%m-%d"))
                grid_from_memory = True

                ce_contract = OptionContract(
                    exchange=ce_c_data["exchange"],
                    trading_symbol=ce_c_data["trading_symbol"],
                    symbol_token=ce_c_data["symbol_token"],
                    expiry=ce_c_data["expiry"],
                    strike=ce_c_data["strike"],
                    option_type=ce_c_data["option_type"],
                    delta=ce_c_data["delta"]
                )
                
                pe_contract = OptionContract(
                    exchange=pe_c_data["exchange"],
                    trading_symbol=pe_c_data["trading_symbol"],
                    symbol_token=pe_c_data["symbol_token"],
                    expiry=pe_c_data["expiry"],
                    strike=pe_c_data["strike"],
                    option_type=pe_c_data["option_type"],
                    delta=pe_c_data["delta"]
                )
                
                # Reconstruct grid
                grid_data = mem["grid"]
                ce_leg_data = grid_data["ce_leg"]
                pe_leg_data = grid_data["pe_leg"]
                
                ce_leg = MasterGridLeg(
                    option_type=ce_leg_data.get("option_type", "CE"),
                    strike=ce_leg_data["strike"],
                    ltp=ce_leg_data["ltp"],
                    delta=ce_leg_data["delta"],
                    target_epm=ce_leg_data["target_epm"],
                    epm_lower_range=ce_leg_data["epm_lower_range"],
                    sl_auto=ce_leg_data["sl_auto"],
                    practical_target=ce_leg_data["practical_target"],
                    expiry=ce_leg_data.get("expiry", getattr(ce_contract, "expiry", "")),
                    trading_symbol=ce_leg_data.get("trading_symbol", getattr(ce_contract, "trading_symbol", ""))
                )
                
                pe_leg = MasterGridLeg(
                    option_type=pe_leg_data.get("option_type", "PE"),
                    strike=pe_leg_data["strike"],
                    ltp=pe_leg_data["ltp"],
                    delta=pe_leg_data["delta"],
                    target_epm=pe_leg_data["target_epm"],
                    epm_lower_range=pe_leg_data["epm_lower_range"],
                    sl_auto=pe_leg_data["sl_auto"],
                    practical_target=pe_leg_data["practical_target"],
                    expiry=pe_leg_data.get("expiry", getattr(pe_contract, "expiry", "")),
                    trading_symbol=pe_leg_data.get("trading_symbol", getattr(pe_contract, "trading_symbol", ""))
                )

                ce_legs = []
                if "ce_legs" in grid_data and isinstance(grid_data["ce_legs"], list):
                    for l_data in grid_data["ce_legs"]:
                        ce_legs.append(MasterGridLeg(
                            option_type=l_data.get("option_type", "CE"),
                            strike=l_data["strike"],
                            ltp=l_data["ltp"],
                            delta=l_data["delta"],
                            target_epm=l_data["target_epm"],
                            epm_lower_range=l_data["epm_lower_range"],
                            sl_auto=l_data["sl_auto"],
                            practical_target=l_data["practical_target"],
                            expiry=l_data.get("expiry", getattr(ce_contract, "expiry", "")),
                            trading_symbol=l_data.get("trading_symbol", getattr(ce_contract, "trading_symbol", ""))
                        ))
                else:
                    ce_legs = [ce_leg]

                pe_legs = []
                if "pe_legs" in grid_data and isinstance(grid_data["pe_legs"], list):
                    for l_data in grid_data["pe_legs"]:
                        pe_legs.append(MasterGridLeg(
                            option_type=l_data.get("option_type", "PE"),
                            strike=l_data["strike"],
                            ltp=l_data["ltp"],
                            delta=l_data["delta"],
                            target_epm=l_data["target_epm"],
                            epm_lower_range=l_data["epm_lower_range"],
                            sl_auto=l_data["sl_auto"],
                            practical_target=l_data["practical_target"],
                            expiry=l_data.get("expiry", getattr(pe_contract, "expiry", "")),
                            trading_symbol=l_data.get("trading_symbol", getattr(pe_contract, "trading_symbol", ""))
                        ))
                else:
                    pe_legs = [pe_leg]

                unique_ce = set(l.strike for l in ce_legs)
                unique_pe = set(l.strike for l in pe_legs)
                if len(unique_ce) < 3 or len(unique_pe) < 3:
                    raise ValueError("Legacy or non-unique strike memory format found. Triggering fresh 3-ITM strike grid calculation.")

                grid = EPMMasterGrid(
                    spot=grid_data["spot"],
                    vix=grid_data["vix"],
                    dte=grid_data["dte"],
                    dte_sqrt=math.sqrt(grid_data["dte"] / 365.0),
                    index_move=grid_data["index_move"],
                    noise_10=grid_data["index_move"] * 0.10,
                    lower_index=grid_data["spot"] - grid_data["index_move"],
                    upper_index=grid_data["spot"] + grid_data["index_move"],
                    ce_leg=ce_leg,
                    pe_leg=pe_leg,
                    ce_legs=ce_legs,
                    pe_legs=pe_legs
                )
                
                spot_price = grid.spot
                vix_val = grid.vix
                dte_days = grid.dte
                ce_ltp = ce_leg.ltp
                pe_ltp = pe_leg.ltp

                # Restore active trade variables if any
                saved_bot_state = mem.get("bot_state", "IDLE")
                if saved_bot_state in ("CE_LONG", "PE_LONG"):
                    bot_state = saved_bot_state
                    active_entry_price = mem.get("active_entry_price", 0.0)
                    active_sl = mem.get("active_sl", 0.0)
                    active_target = mem.get("active_target", 0.0)
                    peak_price = mem.get("peak_price", 0.0)
                    trailing_active = mem.get("trailing_active", False)
                    offloaded = mem.get("offloaded", False)
                    lot_size = mem.get("lot_size", 1)
                    
                    saved_entry_time = mem.get("entry_time")
                    if saved_entry_time:
                        entry_time = datetime.fromisoformat(saved_entry_time)
                    
                    if bot_state == "CE_LONG":
                        active_contract = ce_contract
                    else:
                        active_contract = pe_contract
                    
                    logger.info("⚡ [RECALL ACTIVE POSITION] Resumed active %s trade from memory (Entry: ₹%.2f, SL: ₹%.2f, Target: ₹%.2f)", bot_state, active_entry_price, active_sl, active_target)

            except (KeyError, ValueError, TypeError) as exc:
                logger.warning("⚠️ Saved memory schema mismatch. Resetting and calculating master grid cleanly: %s", exc)
                grid = None
                grid_from_memory = False

    if grid is None:
        logger.info("🆕 [MASTER GRID] Calculating a new Master Grid (3 ITM Strikes for CE & PE) for slot %s...", current_slot)
        grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(smart_api, current_slot, spot_price, spot_open, vix_val)
        ce_contract = ce_contracts[0]
        pe_contract = pe_contracts[0]
        ce_ltp = grid.ce_leg.ltp
        pe_ltp = grid.pe_leg.ltp
        dte_days = grid.dte
        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)

    if grid_from_memory:
        logger.info("=========================================================================")
        logger.info("SENSEX CLOUD BOT - RECALLED MASTER GRID FROM MEMORY")
        logger.info("Spot: %.2f | VIX: %.2f%% | DTE: %.2f | Move: ±%.2f", spot_price, vix_val, dte_days, grid.index_move)
        for idx, leg in enumerate(grid.ce_legs or [grid.ce_leg], 1):
            exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
            logger.info("CE Strike %d (ITM %d%s): LTP/Open ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                        leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
        for idx, leg in enumerate(grid.pe_legs or [grid.pe_leg], 1):
            exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
            logger.info("PE Strike %d (ITM %d%s): LTP/Open ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                        leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
        logger.info("=========================================================================")
        
        flash_msg = format_grid_notification(grid, "🔄 *RECALLED MASTER GRID FROM MEMORY*", spot_price, spot_open, vix_val, dte_days, current_slot)
        send_mobile_alert(flash_msg)
    else:
        logger.info("=========================================================================")
        logger.info("SENSEX CLOUD BOT - MASTER GRID INITIALIZED (Slot: %s IST)", current_slot)
        logger.info("Spot LTP: %.2f (Open: %.2f) | VIX: %.2f%% | DTE: %.2f | Move: ±%.2f", spot_price, spot_open, vix_val, dte_days, grid.index_move)
        for idx, leg in enumerate(grid.ce_legs or [grid.ce_leg], 1):
            exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
            logger.info("CE Strike %d (ITM %d%s): Price ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                        leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
        for idx, leg in enumerate(grid.pe_legs or [grid.pe_leg], 1):
            exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
            logger.info("PE Strike %d (ITM %d%s): Price ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                        leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
        logger.info("=========================================================================")

        # Send Notification
        msg = format_grid_notification(grid, "🔔 *SENSEX MASTER GRID INITIALIZED*", spot_price, spot_open, vix_val, dte_days, current_slot)
        send_mobile_alert(msg)

        # Send Commands Cheat Sheet / Tips at 9:15 AM (Safe Markdown formatting)
        cheat_sheet_msg = (
            "📱 *SENSEX BOT COMMANDS CHEAT SHEET*\n\n"
            "Use these keywords during an active trade to manage your position on the go:\n\n"
            "1. *Add Lots:* `ADD LOTS`\n"
            "   Example: `ADD 2` (Adds 2 more lots at Market price, initial SL remains same)\n\n"
            "2. *Set Trailing Buffer:* `BUFFER POINTS`\n"
            "   Example: `BUFFER 10` (Sets trailing stop-loss distance to 10 points)\n\n"
            "3. *Modify Stop-Loss:* `SL PRICE`\n"
            "   Example: `SL 450` (Manually sets Stop Loss to ₹450)\n\n"
            "4. *Safety Stops:* `STOP` / `HALT` / `EXIT` / `CLOSE`\n"
            "   (Exits all active positions immediately at Market price and halts bot)\n\n"
            "💡 *Smart Scaling (Auto-Activated):*\n"
            "• *Surge Target (3x Risk):* Sells major portion, moves remaining runner lot SL to Cost Price.\n"
            "• *Practical Target:* Sells major portion, moves remaining runner lot SL to Peak - 20 (wider trailing room)."
        )
        send_mobile_alert(cheat_sheet_msg)

        # Log to Excel
        excel_tracker.add_signal({
            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
            "spot": spot_price,
            "vix": vix_val,
            "ce_symbol": ce_contract.trading_symbol,
            "ce_ltp": ce_ltp,
            "ce_low": grid.ce_leg.epm_lower_range,
            "ce_upper": grid.ce_leg.target_epm,
            "pe_symbol": pe_contract.trading_symbol,
            "pe_ltp": pe_ltp,
            "pe_low": grid.pe_leg.epm_lower_range,
            "pe_upper": grid.pe_leg.target_epm,
        })

    # Initialize and start background WebSocket feed for zero-lag live feed
    ws_feed = None
    try:
        ws_feed = LiveWSFeed(
            client_code=os.environ["ANGEL_ONE_CLIENT_CODE"],
            feed_token=smart_api.feed_token,
            api_key=os.environ["ANGEL_ONE_API_KEY"],
            auth_token=smart_api.auth_token
        )
        tokens_to_sub = ["99919000"] + [c.symbol_token for c in (ce_contracts if 'ce_contracts' in locals() else [ce_contract])] + [c.symbol_token for c in (pe_contracts if 'pe_contracts' in locals() else [pe_contract])]
        ws_feed.start(tokens_to_sub)
        logger.info("⚡ Background WebSocket Feed initialized.")
    except Exception as e:
        logger.warning("Could not initialize WebSocket Feed, falling back to HTTP: %s", e)

    # Continuous Monitoring Loop
    poll_interval = 1.0
    is_continuous = "--once" not in sys.argv
    execution_mode = "LIVE" if "--live" in sys.argv else "PAPER"

    # State Machine Variables
    bot_state = "IDLE"  # Options: "IDLE", "CE_LONG", "PE_LONG"
    trades_completed = 0
    max_trades_per_day = 2
    active_contract = None
    active_entry_price = 0.0
    active_sl = 0.0
    active_target = 0.0
    entry_time = None
    lot_size = 1
    loop_counter = 0
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"

    # Trailing Stop-Loss Variables
    original_sl_distance = 0.0
    trailing_active = False
    peak_price = 0.0
    trail_buffer = 5.0
    offloaded = False

    # Memory of lowest prices observed in IDLE state for bounce SL tracking
    recent_ce_low = ce_ltp
    recent_pe_low = pe_ltp
    previous_ce_ltp = ce_ltp
    previous_pe_ltp = pe_ltp

    if is_continuous:
        logger.info("🔄 Entering continuous monitoring loop (Mode: %s, Refreshing 1s in-place)...", execution_mode)
        tg_listener = TelegramCommandListener(os.environ.get("TELEGRAM_BOT_TOKEN"))
        try:
            while True:
                time.sleep(poll_interval)
                checked_at = datetime.now(IST)
                loop_counter += 1

                # Check for slot transition (e.g. from 9:15 AM to 9:45 AM or 9:45 AM to 12:15 PM)
                now_slot = get_current_time_slot()
                if now_slot != current_slot:
                    sys.stdout.write("\n")
                    logger.info("⏰ [SLOT TRANSITION] Time slot changed from %s to %s", current_slot, now_slot)
                    current_slot = now_slot

                    # 1. Fetch current Spot & VIX again for the new grid
                    try:
                        spot_res = smart_api.ltpData("BSE", "SENSEX", "99919000")
                        spot_price = float(spot_res["data"]["ltp"]) if isinstance(spot_res, dict) and spot_res.get("data") else 77500.0
                        spot_open = float(spot_res["data"]["open"]) if isinstance(spot_res, dict) and spot_res.get("data") and spot_res["data"].get("open") else spot_price
                    except Exception:
                        pass

                    try:
                        vix_res = smart_api.ltpData("NSE", "INDIA VIX", "99926017")
                        vix_val = float(vix_res["data"]["ltp"]) if isinstance(vix_res, dict) and vix_res.get("data") else 13.5
                    except Exception:
                        pass

                    # 2. Re-calculate new Master Grid & Contracts for the new slot
                    logger.info("🆕 [MASTER GRID] Calculating a new Master Grid (3 ITM Strikes for CE & PE) for transitioned slot %s...", current_slot)
                    grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(smart_api, current_slot, spot_price, spot_open, vix_val)
                    ce_contract = ce_contracts[0]
                    pe_contract = pe_contracts[0]
                    ce_ltp = grid.ce_leg.ltp
                    pe_ltp = grid.pe_leg.ltp
                    dte_days = grid.dte

                    # Save EPM to memory
                    if bot_state in ("CE_LONG", "PE_LONG"):
                        save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time)
                    else:
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)

                    # Send Telegram Notification for new slot's EPM
                    logger.info("=========================================================================")
                    logger.info("SENSEX CLOUD BOT - MASTER GRID TRANSITIONED")
                    logger.info("Spot LTP: %.2f | VIX: %.2f%% | DTE: %.2f | Move: ±%.2f", spot_price, vix_val, dte_days, grid.index_move)
                    for idx, leg in enumerate(grid.ce_legs or [grid.ce_leg], 1):
                        exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
                        logger.info("CE Strike %d (ITM %d%s): LTP/Open ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                                    leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
                    for idx, leg in enumerate(grid.pe_legs or [grid.pe_leg], 1):
                        exp_info = f" | Exp: {leg.expiry}" if leg.expiry else ""
                        logger.info("PE Strike %d (ITM %d%s): LTP/Open ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                                    leg.strike, idx, exp_info, leg.ltp, leg.delta, leg.epm_lower_range, leg.target_epm, leg.sl_auto, leg.practical_target)
                    logger.info("=========================================================================")

                    msg = format_grid_notification(grid, f"🔔 *SENSEX MASTER GRID UPDATED ({current_slot} Slot)*", spot_price, spot_open, vix_val, dte_days)
                    send_mobile_alert(msg)

                    # Log to Excel
                    excel_tracker.add_signal({
                        "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                        "spot": spot_price,
                        "vix": vix_val,
                        "ce_symbol": ce_contract.trading_symbol,
                        "ce_ltp": ce_ltp,
                        "ce_low": grid.ce_leg.epm_lower_range,
                        "ce_upper": grid.ce_leg.target_epm,
                        "pe_symbol": pe_contract.trading_symbol,
                        "pe_ltp": pe_ltp,
                        "pe_low": grid.pe_leg.epm_lower_range,
                        "pe_upper": grid.pe_leg.target_epm,
                    })

                    # Reset recent lows and previous LTP tracking for the new slot contracts (only if IDLE)
                    if bot_state == "IDLE":
                        recent_ce_low = ce_ltp
                        recent_pe_low = pe_ltp
                    previous_ce_ltp = ce_ltp
                    previous_pe_ltp = pe_ltp

                    # Update WebSocket Feed subscription for the new contracts
                    if ws_feed:
                        try:
                            logger.info("🔌 Closing old WebSocket feed and restarting with new slot tokens...")
                            if ws_feed.ws:
                                ws_feed.ws.close()
                        except Exception as e:
                            logger.debug("Failed to close old WebSocket: %s", e)

                        try:
                            ws_feed = LiveWSFeed(
                                client_code=os.environ["ANGEL_ONE_CLIENT_CODE"],
                                feed_token=smart_api.feed_token,
                                api_key=os.environ["ANGEL_ONE_API_KEY"],
                                auth_token=smart_api.auth_token
                            )
                            ws_feed.start(["99919000", ce_contract.symbol_token, pe_contract.symbol_token])
                            logger.info("⚡ Background WebSocket Feed re-initialized for slot %s.", current_slot)
                        except Exception as e:
                            logger.warning("Could not re-initialize WebSocket Feed: %s", e)

                # Graceful Market Close Exit at 3:30 PM IST (15:30 IST)
                if (checked_at.hour == 15 and checked_at.minute >= 30) or (checked_at.hour > 15):
                    sys.stdout.write("\n")
                    logger.info("🕒 [MARKET CLOSE] Current time is after 3:30 PM IST. Shutting down bot gracefully...")
                    send_mobile_alert("🕒 *MARKET CLOSE REACHED*\nCurrent time is after 3:30 PM IST. Shutting down bot gracefully.")
                    break

                # Check Telegram for remote commands ('LIVE [LOTS]', 'DEMO', 'STOP', 'ADD', 'BUFFER', 'SL')
                cmd, remote_lots = tg_listener.get_new_command()
                if cmd == "STOP":
                    sys.stdout.write("\n")
                    logger.info("🛑 Remote STOP command received via Telegram! Halting execution...")
                    send_mobile_alert("🛑 *REMOTE STOP COMMAND RECEIVED*\nBot execution halted safely.")
                    break
                elif cmd == "LIVE":
                    if remote_lots:
                        lot_size = remote_lots
                    if execution_mode != "LIVE":
                        execution_mode = "LIVE"
                        logger.info("⚠️ [MODE SWITCH] Switched to REAL LIVE TRADING MODE via Telegram (Lot Size: %d).", lot_size)
                        send_mobile_alert(f"🚨 *MODE SWITCHED TO REAL LIVE TRADING*\nLot Size: *{lot_size} Lot(s)* ({lot_size * 20} Qty)\nReal orders will be placed on Angel One.")
                elif cmd == "DEMO" and execution_mode != "PAPER":
                    execution_mode = "PAPER"
                    logger.info("🛡️ [MODE SWITCH] Switched back to SAFE PAPER TRADING MODE via Telegram.")
                    send_mobile_alert("🛡️ *MODE SWITCHED TO PAPER TRADING*\nOrders set to safe demo simulation.")
                elif cmd == "BUFFER":
                    trail_buffer = remote_lots
                    logger.info("⚙️ [TELEGRAM] Trailing buffer manually updated to %.1f points.", trail_buffer)
                    send_mobile_alert(f"⚙️ *TRAILING BUFFER UPDATED*\n\nBuffer manually updated to *{trail_buffer:.1f} points* via Telegram.")
                elif cmd == "SL" and bot_state in ("CE_LONG", "PE_LONG"):
                    active_sl = remote_lots
                    logger.info("⚠️ [TELEGRAM] Stop Loss manually updated to ₹%.2f.", active_sl)
                    send_mobile_alert(f"⚠️ *STOP LOSS UPDATED*\n\nStop Loss manually updated to *₹{active_sl:.2f}* via Telegram.")
                elif cmd == "ADD" and bot_state in ("CE_LONG", "PE_LONG"):
                    lots_to_add = remote_lots
                    qty_to_add = lots_to_add * 20
                    lot_size += lots_to_add
                    logger.info("🚀 [TELEGRAM] Adding %d lot(s) (%d Qty) at market price. New total: %d lots.", lots_to_add, qty_to_add, lot_size)
                    send_mobile_alert(f"🚀 *ADDING LOTS VIA TELEGRAM*\n\n"
                                      f"Adding *{lots_to_add} Lot(s)* ({qty_to_add} Qty) at Market Price.\n"
                                      f"New Total Position: *{lot_size} Lots* ({lot_size * 20} Qty).\n"
                                      f"SL maintained at *₹{active_sl:.2f}* until trailing stop is triggered.")
                    if execution_mode == "LIVE" and active_contract:
                        submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_add)

                # Fetch Live Spot & LTPs (WebSocket with HTTP fallback)
                live_spot = None
                if ws_feed and ws_feed.is_connected:
                    live_spot = ws_feed.prices.get("99919000")
                if live_spot is None:
                    try:
                        live_spot_res = smart_api.ltpData("BSE", "SENSEX", "99919000")
                        live_spot = float(live_spot_res["data"]["ltp"]) if isinstance(live_spot_res, dict) and live_spot_res.get("data") else None
                    except Exception:
                        live_spot = None

                live_ce_ltp = None
                if ws_feed and ws_feed.is_connected:
                    live_ce_ltp = ws_feed.prices.get(ce_contract.symbol_token)
                if live_ce_ltp is None:
                    try:
                        live_ce_res = smart_api.ltpData("BFO", ce_contract.trading_symbol, ce_contract.symbol_token)
                        live_ce_ltp = float(live_ce_res["data"]["ltp"]) if isinstance(live_ce_res, dict) and live_ce_res.get("data") else None
                    except Exception:
                        live_ce_ltp = None

                live_pe_ltp = None
                if ws_feed and ws_feed.is_connected:
                    live_pe_ltp = ws_feed.prices.get(pe_contract.symbol_token)
                if live_pe_ltp is None:
                    try:
                        live_pe_res = smart_api.ltpData("BFO", pe_contract.trading_symbol, pe_contract.symbol_token)
                        live_pe_ltp = float(live_pe_res["data"]["ltp"]) if isinstance(live_pe_res, dict) and live_pe_res.get("data") else None
                    except Exception:
                        live_pe_ltp = None

                if live_spot is None or live_ce_ltp is None or live_pe_ltp is None:
                    logger.warning("⚠️ [API DELAY] Live feed or LTP data returned None (likely Rate Limited). Skipping loop iteration to prevent stale trades.")
                    continue

                # 1. Update Recent Lows while IDLE
                if bot_state == "IDLE":
                    recent_ce_low = min(recent_ce_low, live_ce_ltp)
                    recent_pe_low = min(recent_pe_low, live_pe_ltp)

                # 2. Check for Manual position closure on Broker (LIVE mode only)
                if execution_mode == "LIVE" and bot_state in ("CE_LONG", "PE_LONG") and active_contract:
                    real_qty = check_active_position_qty(smart_api, active_contract.symbol_token)
                    if real_qty == 0:
                        logger.info("🚨 [MANUAL CLOSURE DETECTED] Position for %s has been closed manually on broker. Resetting bot state to IDLE.", active_contract.trading_symbol)
                        send_mobile_alert(f"🚨 *MANUAL POSITION CLOSURE DETECTED*\n"
                                          f"Position for *{active_contract.trading_symbol}* was closed manually on your broker app.\n"
                                          f"Resetting bot state to *IDLE*.")
                        
                        # Log manual exit to Excel Tracker
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": "LIVE",
                            "state": "MANUAL_EXIT",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": live_ce_ltp if bot_state == "CE_LONG" else live_pe_ltp,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = "IDLE"
                        active_contract = None
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        # Reset recent low tracking
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                # 3. State Machine Signal Evaluation
                if bot_state == "IDLE":
                    if trades_completed >= max_trades_per_day:
                        pass
                    else:
                        now_time = datetime.now(IST).time()
                        
                        # EPM Low Levels for CE and PE
                        ce_low_level = grid.ce_leg.epm_lower_range
                        pe_low_level = grid.pe_leg.epm_lower_range
                        
                        # Calculate distances to EPM Low to determine option proximity
                        ce_dist_recent = abs(recent_ce_low - ce_low_level)
                        pe_dist_recent = abs(recent_pe_low - pe_low_level)
                        ce_dist_ltp = abs(live_ce_ltp - ce_low_level)
                        pe_dist_ltp = abs(live_pe_ltp - pe_low_level)
                        
                        ce_min_dist = min(ce_dist_recent, ce_dist_ltp)
                        pe_min_dist = min(pe_dist_recent, pe_dist_ltp)
                        
                        # Symmetric +/- 35 Point Rule: Check if Option Price (recent low or live LTP) is near 35 +/- points of EPM Low
                        ce_in_35_range = (ce_low_level - 35.0 <= recent_ce_low <= ce_low_level + 35.0) or (ce_low_level - 35.0 <= live_ce_ltp <= ce_low_level + 35.0)
                        pe_in_35_range = (pe_low_level - 35.0 <= recent_pe_low <= pe_low_level + 35.0) or (pe_low_level - 35.0 <= live_pe_ltp <= pe_low_level + 35.0)
                        
                        if loop_counter % 30 == 0:
                            logger.debug("[SIGNAL EVAL] Slot: %s | CE dist: %.2f (Near 35pt: %s) | PE dist: %.2f (Near 35pt: %s)",
                                         current_slot, ce_min_dist, ce_in_35_range, pe_min_dist, pe_in_35_range)
                        
                        # First Slot check: Determine which Option Type is nearer to EPM Low to observe first
                        if current_slot == "09:15" and pe_min_dist < ce_min_dist:
                            check_order = ["PE", "CE"]
                        else:
                            check_order = ["CE", "PE"]
                        
                        ce_entry_signal = False
                        active_sl_ce = grid.ce_leg.sl_auto
                        entry_type_str_ce = ""
                        
                        pe_entry_signal = False
                        active_sl_pe = grid.pe_leg.sl_auto
                        entry_type_str_pe = ""
                        
                        for opt_type in check_order:
                            if opt_type == "CE" and not ce_entry_signal and not pe_entry_signal:
                                ce_mfi, ce_prev_mfi, ce_prev_low = get_15m_mfi(smart_api, "BFO", ce_contract.symbol_token, period=5)
                                c_open_15m, c_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", ce_contract.symbol_token)
                                if c_open_15m is None:
                                    c_open_15m = grid.ce_leg.ltp
                                if c_low_15m is None:
                                    c_low_15m = min(recent_ce_low, live_ce_ltp)
                                else:
                                    c_low_15m = min(c_low_15m, recent_ce_low)

                                # Range and MFI Rising Conditions
                                is_in_range_ce = abs(live_ce_ltp - c_open_15m) <= 20.0
                                is_mfi_rising_ce = ce_mfi > ce_prev_mfi

                                # Case 1: Consolidation / Range Breakout Start
                                is_range_move_ce = is_in_range_ce and (live_ce_ltp >= c_open_15m) and is_mfi_rising_ce
                                # Case 2: MFI == 0 and Live LTP reaches Open price after correcting to Low
                                is_mfi_zero_and_bounce_ce = (ce_mfi <= 1.0) and (c_low_15m < c_open_15m) and (live_ce_ltp >= c_open_15m)
                                # Case 3: Retest after unusual/extended jump (>20pts) near +/-10 of swing low
                                is_retest_correction_ce = (not is_in_range_ce) and (ce_prev_low is not None) and (ce_prev_low - 10.0 <= live_ce_ltp <= ce_prev_low + 10.0) and (live_ce_ltp >= c_open_15m) and is_mfi_rising_ce

                                if is_range_move_ce or is_mfi_zero_and_bounce_ce or is_retest_correction_ce:
                                    ce_entry_signal = True
                                    active_sl_ce = max(1.0, c_low_15m - 5.0)
                                    entry_type_str_ce = "Range Breakout Bounce" if is_range_move_ce else ("MFI=0 and Open Bounce" if is_mfi_zero_and_bounce_ce else "Retest Post-Correction Bounce")
                                elif current_slot == "09:15":
                                    if ce_in_35_range:
                                        if live_ce_ltp >= c_open_15m:
                                            ce_entry_signal = True
                                            active_sl_ce = max(1.0, c_low_15m - 5.0)
                                            entry_type_str_ce = "First Slot CE Near EPM Low (+/-35pt) 15m Bounce to Open"
                                else:
                                    if ce_in_35_range and live_ce_ltp >= c_open_15m:
                                        ce_entry_signal = True
                                        active_sl_ce = max(1.0, c_low_15m - 5.0)
                                        entry_type_str_ce = f"{current_slot} Slot CE Near EPM Low (+/-35pt) 15m Bounce to Open"
                                    elif (abs(live_ce_ltp - ce_low_level) <= 2.0) and (live_ce_ltp >= recent_ce_low + 5.0):
                                        ce_entry_signal = True
                                        active_sl_ce = max(1.0, c_low_15m - 5.0)
                                        entry_type_str_ce = f"{current_slot} Slot CE At/Near EPM Low Bounce (+5pt Reversal)"
                                        
                            elif opt_type == "PE" and not ce_entry_signal and not pe_entry_signal:
                                pe_mfi, pe_prev_mfi, pe_prev_low = get_15m_mfi(smart_api, "BFO", pe_contract.symbol_token, period=5)
                                p_open_15m, p_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", pe_contract.symbol_token)
                                if p_open_15m is None:
                                    p_open_15m = grid.pe_leg.ltp
                                if p_low_15m is None:
                                    p_low_15m = min(recent_pe_low, live_pe_ltp)
                                else:
                                    p_low_15m = min(p_low_15m, recent_pe_low)

                                # Range and MFI Rising Conditions
                                is_in_range_pe = abs(live_pe_ltp - p_open_15m) <= 20.0
                                is_mfi_rising_pe = pe_mfi > pe_prev_mfi

                                # Case 1: Consolidation / Range Breakout Start
                                is_range_move_pe = is_in_range_pe and (live_pe_ltp >= p_open_15m) and is_mfi_rising_pe
                                # Case 2: MFI == 0 and Live LTP reaches Open price after correcting to Low
                                is_mfi_zero_and_bounce_pe = (pe_mfi <= 1.0) and (p_low_15m < p_open_15m) and (live_pe_ltp >= p_open_15m)
                                # Case 3: Retest after unusual/extended jump (>20pts) near +/-10 of swing low
                                is_retest_correction_pe = (not is_in_range_pe) and (pe_prev_low is not None) and (pe_prev_low - 10.0 <= live_pe_ltp <= pe_prev_low + 10.0) and (live_pe_ltp >= p_open_15m) and is_mfi_rising_pe

                                if is_range_move_pe or is_mfi_zero_and_bounce_pe or is_retest_correction_pe:
                                    pe_entry_signal = True
                                    active_sl_pe = max(1.0, p_low_15m - 5.0)
                                    entry_type_str_pe = "Range Breakout Bounce" if is_range_move_pe else ("MFI=0 and Open Bounce" if is_mfi_zero_and_bounce_pe else "Retest Post-Correction Bounce")
                                elif current_slot == "09:15":
                                    if pe_in_35_range:
                                        if live_pe_ltp >= p_open_15m:
                                            pe_entry_signal = True
                                            active_sl_pe = max(1.0, p_low_15m - 5.0)
                                            entry_type_str_pe = "First Slot PE Near EPM Low (+/-35pt) 15m Bounce to Open"
                                else:
                                    p_open_15m, p_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", pe_contract.symbol_token)
                                    if p_open_15m is None:
                                        p_open_15m = grid.pe_leg.ltp
                                    if p_low_15m is None:
                                        p_low_15m = min(recent_pe_low, live_pe_ltp)
                                    else:
                                        p_low_15m = min(p_low_15m, recent_pe_low)
                                        
                                    if pe_in_35_range and live_pe_ltp >= p_open_15m:
                                        pe_entry_signal = True
                                        active_sl_pe = max(1.0, p_low_15m - 5.0)
                                        entry_type_str_pe = f"{current_slot} Slot PE Near EPM Low (+/-35pt) 15m Bounce to Open"
                                    elif (abs(live_pe_ltp - pe_low_level) <= 2.0) and (live_pe_ltp >= recent_pe_low + 5.0):
                                        pe_entry_signal = True
                                        active_sl_pe = max(1.0, p_low_15m - 5.0)
                                        entry_type_str_pe = f"{current_slot} Slot PE At/Near EPM Low Bounce (+5pt Reversal)"
                        
                        # Trigger CE Long Entry
                        if ce_entry_signal:
                            bot_state = "CE_LONG"
                            active_contract = ce_contract
                            active_entry_price = live_ce_ltp
                            active_target = grid.ce_leg.practical_target
                            active_sl = active_sl_ce
                            entry_time = datetime.now(IST)
                            qty_to_trade = lot_size * 20
                            original_sl_distance = max(15.0, active_entry_price - active_sl)  # Enforce min 15pt risk distance
                            trailing_active = False
                            peak_price = active_entry_price
                            offloaded = False
                            
                            logger.info("🟢 [ENTRY CE SIGNAL] CE LTP ₹%.2f triggered via %s (SL: ₹%.2f, Target: ₹%.2f)", live_ce_ltp, entry_type_str_ce, active_sl, active_target)
                            send_mobile_alert(f"🟢 *CE ENTRY SIGNAL ALIGNED ({entry_type_str_ce})*\n\n"
                                              f"Contract: *{active_contract.trading_symbol}*\n"
                                              f"Entry Price: ₹{active_entry_price:.2f}\n"
                                              f"Stop Loss: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}\n"
                                              f"Mode: *{execution_mode}* | Lot Size: *{lot_size}* ({qty_to_trade} Qty)")
                            
                            # Place Live Order
                            if execution_mode == "LIVE":
                                submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_trade)
                            
                            # Log to Excel
                            excel_tracker.add_order({
                                "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                                "mode": execution_mode,
                                "state": "ENTRY",
                                "trading_symbol": active_contract.trading_symbol,
                                "price": active_entry_price,
                                "qty": qty_to_trade,
                                "trades_count": trades_completed + 1
                            })
                            # Save state immediately to memory
                            save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time)

                        # Trigger PE Long Entry
                        elif pe_entry_signal:
                            bot_state = "PE_LONG"
                            active_contract = pe_contract
                            active_entry_price = live_pe_ltp
                            active_target = grid.pe_leg.practical_target
                            active_sl = active_sl_pe
                            entry_time = datetime.now(IST)
                            qty_to_trade = lot_size * 20
                            original_sl_distance = max(15.0, active_entry_price - active_sl)  # Enforce min 15pt risk distance
                            trailing_active = False
                            peak_price = active_entry_price
                            offloaded = False
                            
                            logger.info("🟢 [ENTRY PE SIGNAL] PE LTP ₹%.2f triggered via %s (SL: ₹%.2f, Target: ₹%.2f)", live_pe_ltp, entry_type_str_pe, active_sl, active_target)
                            send_mobile_alert(f"🟢 *PE ENTRY SIGNAL ALIGNED ({entry_type_str_pe})*\n\n"
                                              f"Contract: *{active_contract.trading_symbol}*\n"
                                              f"Entry Price: ₹{active_entry_price:.2f}\n"
                                              f"Stop Loss: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}\n"
                                              f"Mode: *{execution_mode}* | Lot Size: *{lot_size}* ({qty_to_trade} Qty)")
                            
                            # Place Live Order
                            if execution_mode == "LIVE":
                                submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_trade)
                            
                            # Log to Excel
                            excel_tracker.add_order({
                                "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                                "mode": execution_mode,
                                "state": "ENTRY",
                                "trading_symbol": active_contract.trading_symbol,
                                "price": active_entry_price,
                                "qty": qty_to_trade,
                                "trades_count": trades_completed + 1
                            })
                            # Save state immediately to memory
                            save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time)

                elif bot_state == "CE_LONG":
                    # --- CE EXIT EVALUATION ---
                    # Calculate 3X risk-reward Take Profit target based on original risk distance
                    surge_target_price = active_entry_price + (3 * original_sl_distance)
                    
                    elapsed_mins = (datetime.now(IST) - entry_time).total_seconds() / 60.0
                    is_surge_window = (elapsed_mins <= 60.0)  # Within 1-2 30-min candles (60 mins)
                    
                    # 1. Check for Surge/Target Trailing SL activation and Smart Offloading
                    is_surge_triggered = is_surge_window and (live_ce_ltp >= surge_target_price)
                    is_target_triggered = (live_ce_ltp >= active_target)
                    is_trailing_triggered = is_surge_triggered or is_target_triggered
                    
                    if is_trailing_triggered:
                        if not trailing_active:
                            trailing_active = True
                            peak_price = live_ce_ltp
                            
                            # Smart Scaling Out (Offload major portion if holding multiple lots)
                            if lot_size > 1 and not offloaded:
                                offloaded = True
                                major_portion = lot_size - 1
                                remaining = 1
                                qty_to_offload = major_portion * 20
                                
                                logger.info("🚀 [SMART SCALING] Triggered. Offloading major portion: %d lots at ₹%.2f", major_portion, live_ce_ltp)
                                if execution_mode == "LIVE":
                                    execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, qty_to_offload, live_ce_ltp)
                                
                                # Adjust SL for the remaining 1 lot
                                if is_surge_triggered:
                                    active_sl = active_entry_price  # Cost Price / Break-even
                                    scale_reason = f"Surge Target (3x RR) hit. Offloaded major portion ({major_portion} lots) at ₹{live_ce_ltp:.2f}. Remaining 1 runner lot SL moved to Cost Price ₹{active_entry_price:.2f}."
                                else:
                                    active_sl = max(active_sl, live_ce_ltp - 20.0)  # Wide 20-point TSL
                                    scale_reason = f"Practical Target hit. Offloaded major portion ({major_portion} lots) at ₹{live_ce_ltp:.2f}. Remaining 1 runner lot SL set to wide Trailing SL ₹{active_sl:.2f} (20-point buffer)."
                                    
                                lot_size = remaining  # We only have the 1 runner lot left now
                                send_mobile_alert(f"🚀 *SMART SCALING OUT ACTIVE*\n\n{scale_reason}")
                            else:
                                # Normal single lot trailing stop activation
                                active_sl = max(active_sl, live_ce_ltp - trail_buffer)
                                logger.info("🔥 [TRAILING ACTIVATED] CE peak reached ₹%.2f. Trailing SL activated at ₹%.2f.", peak_price, active_sl)
                                send_mobile_alert(f"🔥 *TRAILING ACTIVATED*\n\n"
                                                  f"Contract: *{active_contract.trading_symbol}*\n"
                                                  f"Peak Price: ₹{peak_price:.2f}\n"
                                                  f"Trailing SL: ₹{active_sl:.2f}")
                        elif live_ce_ltp > peak_price:
                            peak_price = live_ce_ltp
                            if offloaded:
                                # If scaled out, only practical target remains trailing with wide 20-point stop, surge remains at cost price
                                if not is_surge_triggered:
                                    active_sl = max(active_sl, live_ce_ltp - 20.0)
                                    logger.info("📈 [TRAILING SL RAISED] CE runner lot peak rose to ₹%.2f. Wide TSL: ₹%.2f.", peak_price, active_sl)
                            else:
                                active_sl = max(active_sl, live_ce_ltp - trail_buffer)
                                logger.info("📈 [TRAILING SL RAISED] CE peak rose to ₹%.2f. Trailing SL: ₹%.2f.", peak_price, active_sl)
                    
                    # 2. Stop Loss or Trailing Stop Loss exit
                    if live_ce_ltp <= active_sl:
                        exit_state_str = "EXIT_SL" if not trailing_active else "EXIT_TSL"
                        exit_title_str = "STOP LOSS HIT" if not trailing_active else "TRAILING SL HIT (PROFIT BOOKED!)"
                        
                        logger.info("🔴 [CE EXIT - %s] CE LTP ₹%.2f hit SL ₹%.2f", exit_title_str, live_ce_ltp, active_sl)
                        send_mobile_alert(f"🔴 *CE EXIT - {exit_title_str}*\n\n"
                                          f"Contract: *{active_contract.trading_symbol}*\n"
                                          f"Exit Price: ₹{live_ce_ltp:.2f}\n"
                                          f"SL: ₹{active_sl:.2f} | Practical Target: ₹{active_target:.2f}\n"
                                          f"Trades: {trades_completed + 1}/{max_trades_per_day}")
                        
                        if execution_mode == "LIVE":
                            execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, lot_size * 20, live_ce_ltp)
                        
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": exit_state_str,
                            "trading_symbol": active_contract.trading_symbol,
                            "price": live_ce_ltp,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = "IDLE"
                        active_contract = None
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                    # 3. Standard Practical Target exit or MFI Target Exit (Only if trailing stop is NOT active)
                    c_mfi_val, c_mfi_prev_val, _ = get_15m_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    is_mfi_tp_hit = (c_mfi_val >= 99.0) or (live_ce_ltp >= active_entry_price + 100.0 and c_mfi_val < c_mfi_prev_val)
                    
                    if not trailing_active and (live_ce_ltp >= active_target or is_mfi_tp_hit):
                        tp_reason = "MFI Target (100 / 100pt+ & Declining)" if is_mfi_tp_hit else f"Practical Target (₹{active_target:.2f})"
                        logger.info("🟢 [CE EXIT - TARGET HIT] CE LTP ₹%.2f hit %s", live_ce_ltp, tp_reason)
                        send_mobile_alert(f"🟢 *CE EXIT - TARGET REACHED*\n\n"
                                          f"Reason: {tp_reason}\n"
                                          f"Contract: *{active_contract.trading_symbol}*\n"
                                          f"Exit Price: ₹{live_ce_ltp:.2f} (Entry: ₹{active_entry_price:.2f})\n"
                                          f"Trades: {trades_completed + 1}/{max_trades_per_day}")
                        
                        if execution_mode == "LIVE":
                            execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, lot_size * 20, live_ce_ltp)
                        
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": "EXIT_TP",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": live_ce_ltp,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = "IDLE"
                        active_contract = None
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                elif bot_state == "PE_LONG":
                    # --- PE EXIT EVALUATION ---
                    # Calculate 3X risk-reward Take Profit target based on original risk distance
                    surge_target_price = active_entry_price + (3 * original_sl_distance)
                    
                    elapsed_mins = (datetime.now(IST) - entry_time).total_seconds() / 60.0
                    is_surge_window = (elapsed_mins <= 60.0)  # Within 1-2 30-min candles (60 mins)
                    
                    # 1. Check for Surge/Target Trailing SL activation and Smart Offloading
                    is_surge_triggered = is_surge_window and (live_pe_ltp >= surge_target_price)
                    is_target_triggered = (live_pe_ltp >= active_target)
                    is_trailing_triggered = is_surge_triggered or is_target_triggered
                    
                    if is_trailing_triggered:
                        if not trailing_active:
                            trailing_active = True
                            peak_price = live_pe_ltp
                            
                            # Smart Scaling Out (Offload major portion if holding multiple lots)
                            if lot_size > 1 and not offloaded:
                                offloaded = True
                                major_portion = lot_size - 1
                                remaining = 1
                                qty_to_offload = major_portion * 20
                                
                                logger.info("🚀 [SMART SCALING] Triggered. Offloading major portion: %d lots at ₹%.2f", major_portion, live_pe_ltp)
                                if execution_mode == "LIVE":
                                    execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, qty_to_offload, live_pe_ltp)
                                
                                # Adjust SL for the remaining 1 lot
                                if is_surge_triggered:
                                    active_sl = active_entry_price  # Cost Price / Break-even
                                    scale_reason = f"Surge Target (3x RR) hit. Offloaded major portion ({major_portion} lots) at ₹{live_pe_ltp:.2f}. Remaining 1 runner lot SL moved to Cost Price ₹{active_entry_price:.2f}."
                                else:
                                    active_sl = max(active_sl, live_pe_ltp - 20.0)  # Wide 20-point TSL
                                    scale_reason = f"Practical Target hit. Offloaded major portion ({major_portion} lots) at ₹{live_pe_ltp:.2f}. Remaining 1 runner lot SL set to wide Trailing SL ₹{active_sl:.2f} (20-point buffer)."
                                    
                                lot_size = remaining  # We only have the 1 runner lot left now
                                send_mobile_alert(f"🚀 *SMART SCALING OUT ACTIVE*\n\n{scale_reason}")
                            else:
                                # Normal single lot trailing stop activation
                                active_sl = max(active_sl, live_pe_ltp - trail_buffer)
                                logger.info("🔥 [TRAILING ACTIVATED] PE peak reached ₹%.2f. Trailing SL activated at ₹%.2f.", peak_price, active_sl)
                                send_mobile_alert(f"🔥 *TRAILING ACTIVATED*\n\n"
                                                  f"Contract: *{active_contract.trading_symbol}*\n"
                                                  f"Peak Price: ₹{peak_price:.2f}\n"
                                                  f"Trailing SL: ₹{active_sl:.2f}")
                        elif live_pe_ltp > peak_price:
                            peak_price = live_pe_ltp
                            if offloaded:
                                # If scaled out, only practical target remains trailing with wide 20-point stop, surge remains at cost price
                                if not is_surge_triggered:
                                    active_sl = max(active_sl, live_pe_ltp - 20.0)
                                    logger.info("📈 [TRAILING SL RAISED] PE runner lot peak rose to ₹%.2f. Wide TSL: ₹%.2f.", peak_price, active_sl)
                            else:
                                active_sl = max(active_sl, live_pe_ltp - trail_buffer)
                                logger.info("📈 [TRAILING SL RAISED] PE peak rose to ₹%.2f. Trailing SL: ₹%.2f.", peak_price, active_sl)
                    
                    # 2. Stop Loss or Trailing Stop Loss exit
                    if live_pe_ltp <= active_sl:
                        exit_state_str = "EXIT_SL" if not trailing_active else "EXIT_TSL"
                        exit_title_str = "STOP LOSS HIT" if not trailing_active else "TRAILING SL HIT (PROFIT BOOKED!)"
                        
                        logger.info("🔴 [PE EXIT - %s] PE LTP ₹%.2f hit SL ₹%.2f", exit_title_str, live_pe_ltp, active_sl)
                        send_mobile_alert(f"🔴 *PE EXIT - {exit_title_str}*\n\n"
                                          f"Contract: *{active_contract.trading_symbol}*\n"
                                          f"Exit Price: ₹{live_pe_ltp:.2f}\n"
                                          f"SL: ₹{active_sl:.2f} | Practical Target: ₹{active_target:.2f}\n"
                                          f"Trades: {trades_completed + 1}/{max_trades_per_day}")
                        
                        if execution_mode == "LIVE":
                            execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, lot_size * 20, live_pe_ltp)
                        
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": exit_state_str,
                            "trading_symbol": active_contract.trading_symbol,
                            "price": live_pe_ltp,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = "IDLE"
                        active_contract = None
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                    # 3. Standard Practical Target exit or MFI Target Exit (Only if trailing stop is NOT active)
                    p_mfi_val, p_mfi_prev_val, _ = get_15m_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    is_mfi_tp_hit_pe = (p_mfi_val >= 99.0) or (live_pe_ltp >= active_entry_price + 100.0 and p_mfi_val < p_mfi_prev_val)
                    
                    if not trailing_active and (live_pe_ltp >= active_target or is_mfi_tp_hit_pe):
                        tp_reason_pe = "MFI Target (100 / 100pt+ & Declining)" if is_mfi_tp_hit_pe else f"Practical Target (₹{active_target:.2f})"
                        logger.info("🟢 [PE EXIT - TARGET HIT] PE LTP ₹%.2f hit %s", live_pe_ltp, tp_reason_pe)
                        send_mobile_alert(f"🟢 *PE EXIT - TARGET REACHED*\n\n"
                                          f"Reason: {tp_reason_pe}\n"
                                          f"Contract: *{active_contract.trading_symbol}*\n"
                                          f"Exit Price: ₹{live_pe_ltp:.2f} (Entry: ₹{active_entry_price:.2f})\n"
                                          f"Trades: {trades_completed + 1}/{max_trades_per_day}")
                        
                        if execution_mode == "LIVE":
                            execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, lot_size * 20, live_pe_ltp)
                        
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": "EXIT_TP",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": live_pe_ltp,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = "IDLE"
                        active_contract = None
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                # 4. Save previous LTPs for sharp bounce tracking
                previous_ce_ltp = live_ce_ltp
                previous_pe_ltp = live_pe_ltp

                # 5. Format Shorter Status Line to prevent console wrapping and fix inline refresh
                # We show the dynamic SL and Target of the current active trade in real-time
                if bot_state == "CE_LONG":
                    state_info = f"CE_LONG (SL ₹{active_sl:.1f} / TP ₹{active_target:.1f})"
                elif bot_state == "PE_LONG":
                    state_info = f"PE_LONG (SL ₹{active_sl:.1f} / TP ₹{active_target:.1f})"
                else:
                    state_info = "IDLE"
                    
                status_line = (
                    f"[{checked_at.strftime('%H:%M:%S')}] [{execution_mode}] {state_info} | "
                    f"Trades: {trades_completed}/{max_trades_per_day} | Spot: {live_spot:.2f} | "
                    f"CE: ₹{live_ce_ltp:.2f} | PE: ₹{live_pe_ltp:.2f}"
                )

                if is_github_actions:
                    # Log to stdout only once every 60 seconds on GitHub Actions to prevent log spam
                    if loop_counter % 60 == 1 or loop_counter == 1:
                        logger.info(status_line)
                else:
                    # Inline refresh with carriage return and line clearing ANSI escape sequence for perfect same-line refresh
                    sys.stdout.write(f"\r\033[2K{status_line}")
                    sys.stdout.flush()

        except KeyboardInterrupt:
            sys.stdout.write("\n")
            logger.info("🛑 Monitoring loop stopped by user.")


if __name__ == "__main__":
    try:
        run_cloud_bot()
    except Exception as exc:
        logger.error("❌ CLOUD BOT EXECUTION ERROR: %s", exc, exc_info=True)
        sys.exit(1)

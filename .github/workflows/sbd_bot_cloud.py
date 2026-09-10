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

import contextlib
import io
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
# SYSTEM CONFIGURATION & STRATEGY FEATURE TOGGLES
# =========================================================================
ENABLE_IN_FLIGHT_HANDOVER: bool = True  # Dynamic in-flight strategy switching toggle (Default: True)
MIN_EV_HANDOVER_DELTA: float = 4.0      # Minimum EV delta improvement (+4.0 pts) required to approve handover

# Import Modular 3-Strategy Suite with Dynamic In-Flight Handover
try:
    from state_machine_handover_engine import (
        BaseStrategy,
        EntrySignal,
        ExitSignal,
        HandoverDecision,
        Position,
        PositionSide,
        MarketContext,
        InitialDualMFILowerBandBounceStrategy,
        PreviousHighBreakoutMomentumStrategy,
        DualMFI30mSecondaryReversalStrategy,
        MFITrendReentryMBStrategy,
        PostSLRecoveryReentryStrategy,
        PostBreakdownOversoldBounceStrategy,
        DynamicSwingLowBreakoutRetestStrategy,
        StateMachineHandoverEngine,
    )
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from state_machine_handover_engine import (
        BaseStrategy,
        EntrySignal,
        ExitSignal,
        HandoverDecision,
        Position,
        PositionSide,
        MarketContext,
        InitialDualMFILowerBandBounceStrategy,
        PreviousHighBreakoutMomentumStrategy,
        DualMFI30mSecondaryReversalStrategy,
        MFITrendReentryMBStrategy,
        PostSLRecoveryReentryStrategy,
        PostBreakdownOversoldBounceStrategy,
        DynamicSwingLowBreakoutRetestStrategy,
        StateMachineHandoverEngine,
    )


def map_entry_type_to_strategy_name(entry_type_str: str) -> str:
    """Map human-readable entry signal descriptions to standardized Strategy names in Modular 3-Strategy Suite."""
    if not entry_type_str:
        return "Previous High Breakout Momentum Entry"
    if "Recovery" in entry_type_str or "Post-SL" in entry_type_str:
        return "One-Time Post-SL Recovery Re-Entry (+2 Lots)"
    elif "Swing Low" in entry_type_str or "SWING" in entry_type_str or "Retest" in entry_type_str:
        return "Dynamic Swing Low First Breakout Retest Entry"
    elif "Breakout" in entry_type_str or "Previous High" in entry_type_str:
        return "Previous High Breakout Momentum Entry"
    elif "Post-Breakdown" in entry_type_str:
        return "Post-Breakdown Oversold Bounce Entry"
    elif "MB Consolidation" in entry_type_str or ("Re-Entry" in entry_type_str and "Recovery" not in entry_type_str):
        return "MFI(14) Trend Re-Entry (MB Consolidation)"
    elif "Secondary Reversal" in entry_type_str or "30m Dual MFI" in entry_type_str:
        return "30m Dual MFI Secondary Reversal Option"
    else:
        return "Previous High Breakout Momentum Entry"


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
                            elif first_word in ("Y", "YES", "/Y"):
                                return "Y", None
                            elif first_word in ("N", "NO", "/N"):
                                return "N", None
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
                                        points = max(0.01, float(parts[1]))
                                    except ValueError:
                                        pass
                                return "BUFFER", points
                            elif first_word in ("LIMIT", "MAX_TRADES", "/LIMIT"):
                                max_t = 2
                                if len(parts) > 1 and parts[1].isdigit():
                                    max_t = max(1, int(parts[1]))
                                return "LIMIT", max_t
                            elif first_word == "SL":
                                price = 0.0
                                if len(parts) > 1:
                                    try:
                                        price = max(1.0, float(parts[1]))
                                    except ValueError:
                                        pass
                                return "SL", price
                            elif first_word == "REENTRY":
                                sub = parts[1] if len(parts) > 1 else "ON"
                                if sub in ("ON", "TRUE", "ENABLE", "1"):
                                    return "REENTRY ON", None
                                else:
                                    return "REENTRY OFF", None
        except Exception:
            pass
        return None, None


def parse_angel_order_response(res: Any) -> tuple[str | None, str]:
    """Parse Angel One SmartAPI placeOrder response safely. Returns (order_id, error_message)."""
    if not res:
        return None, "Empty response from SmartAPI"
    if isinstance(res, str) and res.strip():
        return res.strip(), ""
    if isinstance(res, dict):
        status = res.get("status")
        if status is True and res.get("data"):
            data = res["data"]
            if isinstance(data, dict):
                oid = str(data.get("orderid") or data.get("order_id") or "").strip()
                if oid:
                    return oid, ""
            elif isinstance(data, str) and data.strip():
                return data.strip(), ""
        err_msg = str(res.get("message") or res.get("errorcode") or res)
        return None, err_msg
    return None, str(res)


def reauthenticate_smartapi(smart_api: Any) -> bool:
    """Silently re-authenticates and refreshes JWT & Feed tokens if session expired or collided."""
    try:
        import pyotp
        import requests
        req_keys = ("ANGEL_ONE_API_KEY", "ANGEL_ONE_CLIENT_CODE", "ANGEL_ONE_PASSWORD", "ANGEL_ONE_TOTP_SECRET")
        if any(not os.environ.get(k) for k in req_keys):
            return False
        try:
            pub_ip = requests.get("https://api.ipify.org", timeout=3.0).text.strip()
        except Exception:
            pub_ip = "117.97.214.136"
        smart_api.clientPublicIP = pub_ip
        login_response = smart_api.generateSession(
            os.environ["ANGEL_ONE_CLIENT_CODE"],
            os.environ["ANGEL_ONE_PASSWORD"],
            pyotp.TOTP(os.environ["ANGEL_ONE_TOTP_SECRET"]).now(),
        )
        if isinstance(login_response, dict) and login_response.get("status") is True and login_response.get("data"):
            smart_api.feed_token = login_response["data"].get("feedToken")
            smart_api.auth_token = login_response["data"].get("jwtToken")
            logger.info("🔄 [AUTH RECOVERY] Successfully re-authenticated SmartAPI session in-flight.")
            return True
    except Exception as exc:
        logger.warning("⚠️ Re-authentication attempt failed: %s", exc)
    return False


def submit_angel_order(smart_api: Any, trading_symbol: str, symbol_token: str, transaction_type: str = "BUY", quantity: int = 10) -> Any:
    """Submit real Market Order to Angel One SmartAPI with product type fallback, auto-reauth, and robust error handling."""
    qty_val = max(1, int(quantity))
    for attempt in range(1, 3):
        for product_type in ("INTRADAY", "CARRYFORWARD"):
            try:
                order_params = {
                    "variety": "NORMAL",
                    "tradingsymbol": str(trading_symbol).strip(),
                    "symboltoken": str(symbol_token).strip(),
                    "transactiontype": transaction_type.upper(),
                    "exchange": "BFO",
                    "ordertype": "MARKET",
                    "producttype": product_type,
                    "duration": "DAY",
                    "price": "0",
                    "squareoff": "0",
                    "stoploss": "0",
                    "quantity": str(qty_val),
                }
                res = smart_api.placeOrder(order_params)
                order_id, err_msg = parse_angel_order_response(res)
                if order_id:
                    logger.info("⚡ [REAL ORDER SUBMITTED] %s %d %s (%s) | Order ID: %s", transaction_type, qty_val, trading_symbol, product_type, order_id)
                    send_mobile_alert(f"🚨 *REAL ORDER PLACED ON ANGEL ONE*\n\nAction: *{transaction_type}*\nContract: *{trading_symbol}*\nQuantity: *{qty_val}*\nOrder ID: `{order_id}`")
                    return order_id
                else:
                    logger.warning("⚠️ SmartAPI Order rejected with producttype=%s: %s", product_type, err_msg)
                    err_lower = err_msg.lower()
                    if any(kw in err_lower for kw in ["token", "session", "unauthorized", "expired", "ag8001", "login", "auth"]):
                        logger.info("🔄 Session token error detected during order placement. Triggering instant re-auth...")
                        reauthenticate_smartapi(smart_api)
                        break  # Retry loop with refreshed credentials
            except Exception as exc:
                exc_str = str(exc)
                logger.warning("⚠️ Exception submitting order with producttype=%s: %s", product_type, exc)
                if any(kw in exc_str.lower() for kw in ["token", "session", "unauthorized", "expired", "ag8001", "login", "auth"]):
                    reauthenticate_smartapi(smart_api)
                    break
    
    logger.error("❌ Real Order Submission Failed for %s %d %s", transaction_type, qty_val, trading_symbol)
    send_mobile_alert(f"⚠️ *ORDER SUBMISSION ERROR*\nFailed to place {transaction_type} for {trading_symbol}. Check Angel One account permissions.")
    return None


def send_telegram_voice_alert(message: str) -> None:
    """Send voice alert as audio clip to Telegram if gtts is available."""
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not bot_token or not chat_id:
        return
        
    msg_upper = message.upper()
    is_event = ("ENTRY" in msg_upper or "EXIT" in msg_upper or "TARGET" in msg_upper or "STOP LOSS" in msg_upper or "TSL" in msg_upper or "SL HIT" in msg_upper or "SIGNAL" in msg_upper or "MFI" in msg_upper or "GRID" in msg_upper or "SWAP" in msg_upper or "BREAKOUT" in msg_upper)
    if not is_event:
        return
        
    try:
        clean_text = message.replace("*", "").replace("`", "").replace("🟢", "").replace("🔴", "").replace("🔥", "").replace("SBD_bot", "")
        lines = [line.strip() for line in clean_text.split("\n") if line.strip()][:3]
        voice_text = " . ".join(lines)
        
        try:
            from gtts import gTTS
        except ImportError:
            logger.warning("⚠️ gTTS module not installed. Run 'pip install gtts' to enable Telegram Voice Alerts.")
            return

        import io
        tts = gTTS(text=voice_text, lang='en', slow=False)
        fp = io.BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        
        import urllib.request
        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        url = f"https://api.telegram.org/bot{bot_token}/sendVoice"
        
        body = []
        body.append(f'--{boundary}'.encode('utf-8'))
        body.append(f'Content-Disposition: form-data; name="chat_id"'.encode('utf-8'))
        body.append(''.encode('utf-8'))
        body.append(str(chat_id).encode('utf-8'))
        body.append(f'--{boundary}'.encode('utf-8'))
        body.append(f'Content-Disposition: form-data; name="voice"; filename="alert.ogg"'.encode('utf-8'))
        body.append('Content-Type: audio/ogg'.encode('utf-8'))
        body.append(''.encode('utf-8'))
        body.append(fp.read())
        body.append(f'--{boundary}--'.encode('utf-8'))
        body.append(''.encode('utf-8'))
        
        req_body = b'\r\n'.join(body)
        
        req = urllib.request.Request(url, data=req_body)
        req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
        req.add_header('User-Agent', 'Mozilla/5.0')
        
        with urllib.request.urlopen(req, timeout=8) as response:
            if response.status == 200:
                logger.info("🎙️ Telegram Voice Alert sent successfully.")
    except Exception as e:
        logger.warning("Failed to send Telegram Voice Alert: %s", e)


def send_mobile_alert(message: str) -> None:
    """Send mobile push notifications via Telegram Bot API or CallMeBot WhatsApp API with automatic fallback."""
    # Send Voice Alert to Telegram if applicable
    send_telegram_voice_alert(message)

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
            logger.warning("Failed to send Telegram Markdown alert, retrying plain text: %s", e)
            try:
                # Retry without Markdown formatting so syntax errors never prevent notification delivery
                clean_msg = message.replace("*", "").replace("`", "")
                payload_plain = urllib.parse.urlencode({"chat_id": chat_id, "text": clean_msg}).encode("utf-8")
                req_plain = urllib.request.Request(url, data=payload_plain, headers={"Content-Type": "application/x-www-form-urlencoded"})
                with urllib.request.urlopen(req_plain, timeout=5) as resp_plain:
                    if resp_plain.status == 200:
                        logger.info("📱 Telegram plain text notification sent successfully.")
            except Exception as e2:
                logger.warning("Failed to send Telegram plain text alert: %s", e2)

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

    # Infer strike step from unique contracts
    strikes = sorted([c.strike for c in unique_contracts if c.strike > 0])
    diffs = [strikes[i + 1] - strikes[i] for i in range(len(strikes) - 1)]
    positive_diffs = [d for d in diffs if d > 0]
    strike_step = min(positive_diffs) if positive_diffs else 100.0

    atm_strike = round(spot / strike_step) * strike_step

    if option_type == "CE":
        # Strictly ITM for CE: strike < spot (e.g. for spot 76810: 76700, 76600, 76500)
        itm = sorted([c for c in unique_contracts if c.strike < spot], key=lambda c: c.strike, reverse=True)
        if not itm:
            itm = sorted([c for c in unique_contracts if c.strike <= spot + 100.0], key=lambda c: c.strike, reverse=True)
        return itm[:count]
    else:
        # Strictly ITM for PE: strike > spot (e.g. for spot 76810: 76900, 77000, 77100)
        itm = sorted([c for c in unique_contracts if c.strike > spot], key=lambda c: c.strike)
        if not itm:
            itm = sorted([c for c in unique_contracts if c.strike >= spot - 100.0], key=lambda c: c.strike)
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
    import requests
    from SmartApi import SmartConnect

    try:
        pub_ip = requests.get("https://api.ipify.org", timeout=3.0).text.strip()
    except Exception:
        pub_ip = "117.99.43.62"

    smart_api = SmartConnect(api_key=os.environ["ANGEL_ONE_API_KEY"])
    smart_api.clientPublicIP = pub_ip
    smart_api.clientLocalIP = "127.0.0.1"
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


_last_api_call_time = 0.0

def safe_get_candle_data(smart_api: Any, params: dict, max_retries: int = 5) -> dict | None:
    """Fetch candle data from SmartAPI with global rate limiting and exponential backoff retry.
    Prevents AB1021 'Too many requests' errors.
    """
    global _last_api_call_time
    if smart_api is None:
        return None

    fetch_fn = getattr(smart_api, "getCandleData", getattr(smart_api, "getCandle", None))
    if not callable(fetch_fn):
        return None

    # Minimum spacing between API calls to stay well within Angel One rate limits (3 req/sec)
    min_interval = 0.35  # seconds

    for attempt in range(1, max_retries + 1):
        now = time.time()
        elapsed = now - _last_api_call_time
        if elapsed < min_interval:
            time.sleep(min_interval - elapsed)
        _last_api_call_time = time.time()

        try:
            res = fetch_fn(params)
            if isinstance(res, dict):
                error_code = str(res.get("errorcode") or "")
                message = str(res.get("message") or "").lower()
                
                if "ab1021" in error_code.lower() or "too many requests" in message or "rate limit" in message:
                    backoff = 0.5 * (2 ** attempt)
                    logger.warning("⚠️ SmartAPI rate limit hit (AB1021) on attempt %d/%d. Sleeping %.2fs before retry...", attempt, max_retries, backoff)
                    time.sleep(backoff)
                    continue
                
                if res.get("status") is True and res.get("data"):
                    return res
                
                if attempt < max_retries:
                    time.sleep(0.3 * attempt)
                    continue
        except Exception as exc:
            exc_str = str(exc).lower()
            if "ab1021" in exc_str or "too many requests" in exc_str:
                backoff = 0.5 * (2 ** attempt)
                logger.warning("⚠️ SmartAPI exception (Too many requests) on attempt %d/%d. Sleeping %.2fs...", attempt, max_retries, backoff)
                time.sleep(backoff)
            else:
                logger.debug("Exception in safe_get_candle_data (attempt %d/%d): %s", attempt, max_retries, exc)
                time.sleep(0.3 * attempt)

    return None


def calculate_atr_and_stddev(smart_api: Any, exchange: str, symbol_token: str) -> tuple[float, float]:
    """Calculate ATR(14) and StdDev(20) dynamically from 15-minute candles."""
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=5)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= 20:
                highs = [float(c[2]) for c in candles]
                lows = [float(c[3]) for c in candles]
                closes = [float(c[4]) for c in candles]
                
                # ATR(14)
                tr_list = []
                for i in range(1, len(candles)):
                    tr = max(
                        highs[i] - lows[i],
                        abs(highs[i] - closes[i - 1]),
                        abs(lows[i] - closes[i - 1])
                    )
                    tr_list.append(tr)
                atr_14 = sum(tr_list[-14:]) / 14.0 if len(tr_list) >= 14 else 20.0

                # StdDev(20)
                sub_closes = closes[-20:]
                mean_c = sum(sub_closes) / 20.0
                var_c = sum((x - mean_c) ** 2 for x in sub_closes) / 20.0
                std_dev_20 = math.sqrt(var_c)

                return atr_14, std_dev_20
    except Exception as exc:
        logger.debug("Error calculating ATR and StdDev: %s", exc)
    return 20.0, 15.0


def load_grid_state_epm_low(option_type: str = "CE") -> float | None:
    """Load the EPM Low from last saved grid state file (grid_state.json or bot_state_memory.json).
    No hardcoding.
    """
    import json
    for filename in ("grid_state.json", "bot_state_memory.json"):
        path = Path(filename)
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # Direct top-level fields in grid_state.json
                if option_type.upper() == "CE" and "ce_epm_low" in data and data["ce_epm_low"] is not None:
                    return float(data["ce_epm_low"])
                if option_type.upper() == "PE" and "pe_epm_low" in data and data["pe_epm_low"] is not None:
                    return float(data["pe_epm_low"])
                
                # Nested grid object in bot_state_memory.json
                grid_data = data.get("grid") or {}
                leg_key = "ce_leg" if option_type.upper() == "CE" else "pe_leg"
                if leg_key in grid_data and "epm_lower_range" in grid_data[leg_key]:
                    return float(grid_data[leg_key]["epm_lower_range"])
            except Exception as e:
                logger.debug("Error reading %s for EPM low: %s", filename, e)
    return None


def get_current_15m_candle_ohl(smart_api: Any, exchange: str, symbol_token: str) -> tuple[float | None, float | None]:
    """Fetch the current 15-minute candle's Open and Low prices from SmartAPI.
    Returns (open, low) on success, or (None, None) on error.
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=3)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
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
        from_dt = now_dt - timedelta(days=7)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
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


def get_1h_mfi(smart_api: Any, exchange: str, symbol_token: str, period: int = 5) -> tuple[float, float]:
    """Calculate Money Flow Index (MFI) on 1-hour timeframe.
    Returns (curr_mfi_1h, prev_mfi_1h).
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=10)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "ONE_HOUR",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= period + 1:
                typical_prices = []
                volumes = []
                for c in candles:
                    if len(c) >= 6:
                        h, l, cl, v = float(c[2]), float(c[3]), float(c[4]), float(c[5])
                        typical_prices.append((h + l + cl) / 3.0)
                        volumes.append(v if v > 0 else 1.0)

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
                    return curr_mfi, prev_mfi
    except Exception as e:
        logger.debug("Error calculating 1h MFI: %s", e)
    return 50.0, 50.0


def get_weekly_open_price(smart_api: Any) -> float:
    """Fetch SENSEX Index Weekly Open price (Monday 09:15 AM Open)."""
    try:
        now_dt = datetime.now(IST)
        days_since_monday = now_dt.weekday()  # Monday = 0
        monday_dt = now_dt - timedelta(days=days_since_monday)
        from_str = monday_dt.strftime("%Y-%m-%d 09:15")
        to_str = now_dt.strftime("%Y-%m-%d %H:%M")
        
        params = {
            "exchange": "BSE",
            "symboltoken": "99919000",
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_str,
            "todate": to_str
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if candles and len(candles[0]) >= 2:
                weekly_open = float(candles[0][1])  # First 15m candle open on Monday
                return weekly_open
    except Exception as e:
        logger.debug("Error fetching Weekly Open: %s", e)
    return 77000.0


def get_mfi_multi_period(smart_api: Any, exchange: str, symbol_token: str, timeframe: str = "FIFTEEN_MINUTE", periods: list[int] = [5, 14], return_extra: bool = False) -> Any:
    """Calculate Money Flow Index for multiple periods (e.g. 5 and 14) on specified timeframe.
    Returns (curr_mfis, prev_mfis) or (curr_mfis, prev_mfis, prev_prev_mfis, extra_data) if return_extra is True.
    """
    max_period = max(periods)
    curr_mfis = {p: 50.0 for p in periods}
    prev_mfis = {p: 50.0 for p in periods}
    prev_prev_mfis = {p: 50.0 for p in periods}
    extra_data = {
        "prev_close": 0.0,
        "prev_prev_low": 0.0,
    }
    
    try:
        now_dt = datetime.now(IST)
        lookback_days = 10 if timeframe in ("THIRTY_MINUTE", "ONE_HOUR") else (1 if timeframe == "ONE_MINUTE" else 7)
        from_dt = now_dt - timedelta(days=lookback_days)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": timeframe,
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= 3:
                extra_data["prev_close"] = float(candles[-2][4])
                extra_data["prev_prev_low"] = float(candles[-3][3])
            elif len(candles) == 2:
                extra_data["prev_close"] = float(candles[-1][4])
                extra_data["prev_prev_low"] = float(candles[-2][3])
            elif len(candles) == 1:
                extra_data["prev_close"] = float(candles[-1][4])
                extra_data["prev_prev_low"] = float(candles[-1][3])
                
            if len(candles) >= max_period + 1:
                typical_prices = []
                volumes = []
                for c in candles:
                    if len(c) >= 6:
                        h, l, cl, v = float(c[2]), float(c[3]), float(c[4]), float(c[5])
                        typical_prices.append((h + l + cl) / 3.0)
                        volumes.append(v if v > 0 else 1.0)

                for p in periods:
                    def calc_mfi_at(end_idx, period_val):
                        pos_mf = 0.0
                        neg_mf = 0.0
                        for i in range(end_idx - period_val + 1, end_idx + 1):
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

                    if len(typical_prices) >= p + 3:
                        curr_mfis[p] = calc_mfi_at(len(typical_prices) - 1, p)
                        prev_mfis[p] = calc_mfi_at(len(typical_prices) - 2, p)
                        prev_prev_mfis[p] = calc_mfi_at(len(typical_prices) - 3, p)
                    elif len(typical_prices) >= p + 2:
                        curr_mfis[p] = calc_mfi_at(len(typical_prices) - 1, p)
                        prev_mfis[p] = calc_mfi_at(len(typical_prices) - 2, p)
                        prev_prev_mfis[p] = prev_mfis[p]
                    elif len(typical_prices) >= p + 1:
                        curr_mfis[p] = calc_mfi_at(len(typical_prices) - 1, p)
                        prev_mfis[p] = curr_mfis[p]
                        prev_prev_mfis[p] = curr_mfis[p]
    except Exception as e:
        logger.debug("Error calculating multi-period MFI: %s", e)
        
    if return_extra:
        return curr_mfis, prev_mfis, prev_prev_mfis, extra_data
    return curr_mfis, prev_mfis


def get_3m_bollinger_bands(smart_api: Any, exchange: str, symbol_token: str, period: int = 20, std_dev_mult: float = 2.0) -> tuple[float | None, float | None, float | None, float | None, float | None, float | None, float | None, float | None]:
    """Fetch recent 3-minute candles and calculate Bollinger Bands (20 SMA, +/- 2 StdDev).
    Returns (middle_band, upper_band, lower_band, c_open, c_low, c_close, prev_c_close, swing_low_3m).
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=2)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "THREE_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= period:
                closes = [float(c[4]) for c in candles[-period:]]
                c_open = float(candles[-1][1])
                c_low = float(candles[-1][3])
                c_close = float(candles[-1][4])
                
                prev_c_close = float(candles[-2][4]) if len(candles) >= 2 else c_close
                lows_3m = [float(c[3]) for c in candles[-10:]]
                swing_low_3m = min(lows_3m) if lows_3m else c_low

                mean = sum(closes) / float(period)
                variance = sum((x - mean) ** 2 for x in closes) / float(period)
                std_dev = math.sqrt(variance)

                middle_band = mean
                upper_band = mean + (std_dev_mult * std_dev)
                lower_band = mean - (std_dev_mult * std_dev)

                return middle_band, upper_band, lower_band, c_open, c_low, c_close, prev_c_close, swing_low_3m
    except Exception as e:
        logger.debug("Error calculating 3m Bollinger Bands: %s", e)
    return None, None, None, None, None, None, None, None


def check_green_breakout_structure(smart_api: Any, exchange: str, symbol_token: str, live_price: float) -> tuple[bool, str]:
    """Identify sustainable breakout: consecutive green candles OR a breach of the prior high with a green body close."""
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=2)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIFTEEN_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= 2:
                last = candles[-1]
                prev = candles[-2]
                
                prev_high = float(prev[2])
                prev_open = float(prev[1])
                prev_close = float(prev[4])
                
                last_open = float(last[1])
                
                is_last_green = live_price > last_open
                is_prev_green = prev_close > prev_open
                breaches_prev_high = live_price > prev_high
                
                # 1. Sustainable Breakout Rule: Consecutive green candles OR breach of prior high in Green
                has_consecutive_green = is_last_green and is_prev_green
                has_green_high_breach = breaches_prev_high and is_last_green
                
                if has_consecutive_green or has_green_high_breach:
                    return True, "Sustainable Breakout Confirmed (Consecutive Green or High Breach in Green)"
                else:
                    return False, "Choppy consolidation without consecutive green or prior high breach in Green"
    except Exception as e:
        logger.debug("Error checking breakout structure: %s", e)
    return True, "Default breakout pass"


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


def save_bot_memory_full(
    trades_completed: int,
    grid_time_slot: str,
    grid: EPMMasterGrid,
    ce_contract: OptionContract,
    pe_contract: OptionContract,
    bot_state: str,
    active_contract: OptionContract | None,
    active_entry_price: float,
    active_sl: float,
    active_target: float,
    peak_price: float,
    trailing_active: bool,
    offloaded: bool,
    lot_size: int,
    entry_time: datetime | None,
    entry_mfi_falling_15m: bool = False,
    active_strategy: str = "",
    initial_hard_sl: float = 0.0,
    handover_history: list[dict] | None = None,
) -> None:
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
            "entry_mfi_falling_15m": entry_mfi_falling_15m,
            "active_strategy": active_strategy,
            "initial_hard_sl": initial_hard_sl,
            "handover_history": handover_history or [],
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
        
        # Save exact grid state file as grid_state.json
        grid_state_data = {
            "date": datetime.now(IST).strftime("%Y-%m-%d"),
            "grid_time_slot": grid_time_slot,
            "ce_epm_low": grid.ce_leg.epm_lower_range if grid and grid.ce_leg else None,
            "pe_epm_low": grid.pe_leg.epm_lower_range if grid and grid.pe_leg else None,
            "ce_symbol": ce_contract.trading_symbol if ce_contract else "",
            "pe_symbol": pe_contract.trading_symbol if pe_contract else "",
            "grid": data.get("grid")
        }
        try:
            with open("grid_state.json", "w", encoding="utf-8") as f_grid:
                json.dump(grid_state_data, f_grid, indent=4)
        except Exception as exc_grid:
            logger.warning("Failed to save grid_state.json: %s", exc_grid)

        logger.info("💾 Bot state, grid state, and active trade memory saved successfully.")
    except Exception as e:
        logger.warning("Failed to save bot active memory: %s", e)


def save_bot_memory(trades_completed: int, grid_time_slot: str, grid: EPMMasterGrid, ce_contract: OptionContract, pe_contract: OptionContract) -> None:
    save_bot_memory_full(trades_completed, grid_time_slot, grid, ce_contract, pe_contract, "IDLE", None, 0.0, 0.0, 0.0, 0.0, False, False, 1, None, False)


def get_current_1m_candles(smart_api: Any, exchange: str, symbol_token: str, count_mins: int = 15) -> list:
    """Fetch recent 1-minute candles from SmartAPI.
    Each candle is: [timestamp, open, high, low, close, volume]
    """
    try:
        now_dt = datetime.now(IST)
        from_dt = now_dt - timedelta(days=1)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "ONE_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            return res["data"]
    except Exception as e:
        logger.debug("Error fetching 1m candles: %s", e)
    return []


def check_prior_breakouts(smart_api: Any, token: str) -> bool:
    """Fetch 5-minute candles to check if prior breakouts from the swing low have already happened."""
    candles = get_current_5m_candles(smart_api, "BFO", token)
    if not candles or len(candles) < 2:
        return False
    breakout_candles_count = 0
    for c in candles:
        c_open = float(c[1])
        c_close = float(c[4])
        if c_close - c_open >= 30.0:
            breakout_candles_count += 1
    return breakout_candles_count >= 2


def check_5m_breakout_and_reversal(smart_api: Any, token: str, live_ltp: float) -> tuple[bool, float | None]:
    """Check if there was a 5-minute green breakout candle and price is now testing/bouncing from initial swing low reversing up to open."""
    candles = get_current_5m_candles(smart_api, "BFO", token)
    if not candles or len(candles) < 2:
        return False, None
    
    lows = [float(c[3]) for c in candles]
    swing_low = min(lows)
    
    has_5m_green_breakout = False
    breakout_open_price = None
    for c in candles[-3:]:
        c_open = float(c[1])
        c_close = float(c[4])
        if c_close > c_open + 5.0:
            has_5m_green_breakout = True
            breakout_open_price = c_open
            break
            
    if has_5m_green_breakout and breakout_open_price:
        if (swing_low <= live_ltp <= swing_low + 25.0) or (live_ltp >= breakout_open_price):
            if live_ltp >= breakout_open_price:
                return True, swing_low
    return False, None


def load_delta_map() -> dict[str, dict[str, Any]]:
    """Load official Angel One Script Master metadata map from local files if available."""
    import os, json
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
    buffer: float = 0.13,
) -> tuple[EPMMasterGrid, list[OptionContract], list[OptionContract]]:
    """Fetch option contracts, select top 3 ITM strikes for CE and PE, calculate EPM grid for all of them."""
    with contextlib.redirect_stdout(io.StringIO()):
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
        buffer=buffer,
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
        from_dt = now_dt - timedelta(days=2)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": "FIVE_MINUTE",
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(smart_api, params)
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
    Supports product type fallback (INTRADAY / CARRYFORWARD) and auto session recovery.
    """
    qty_val = max(1, int(quantity))
    
    for attempt in range(1, 3):
        # 1. Try Market Sell Order with product type fallback
        for product_type in ("INTRADAY", "CARRYFORWARD"):
            try:
                order_params = {
                    "variety": "NORMAL",
                    "tradingsymbol": str(trading_symbol).strip(),
                    "symboltoken": str(symbol_token).strip(),
                    "transactiontype": "SELL",
                    "exchange": "BFO",
                    "ordertype": "MARKET",
                    "producttype": product_type,
                    "duration": "DAY",
                    "price": "0",
                    "squareoff": "0",
                    "stoploss": "0",
                    "quantity": str(qty_val),
                }
                res = smart_api.placeOrder(order_params)
                order_id, err_msg = parse_angel_order_response(res)
                if order_id:
                    logger.info("⚡ [MARKET SELL ORDER SUCCESS] (%s) | Order ID: %s", product_type, order_id)
                    send_mobile_alert(f"🔴 *SELL ORDER EXECUTED*\nContract: *{trading_symbol}*\nQty: *{qty_val}*\nOrder ID: `{order_id}`")
                    return order_id
                else:
                    logger.warning("⚠️ Market sell rejected with producttype=%s: %s", product_type, err_msg)
                    err_lower = err_msg.lower()
                    if any(kw in err_lower for kw in ["token", "session", "unauthorized", "expired", "ag8001", "login", "auth"]):
                        logger.info("🔄 Session token error on sell. Triggering instant re-auth...")
                        reauthenticate_smartapi(smart_api)
                        break
            except Exception as exc:
                exc_str = str(exc)
                logger.warning("⚠️ Exception on market sell with producttype=%s: %s", product_type, exc)
                if any(kw in exc_str.lower() for kw in ["token", "session", "unauthorized", "expired", "ag8001", "login", "auth"]):
                    reauthenticate_smartapi(smart_api)
                    break
        
        # 2. Try Failsafe Limit Sell Order (Sell at LTP - 10 points to guarantee execution)
        limit_price = max(2.0, float(ltp) - 10.0)
        limit_price_str = f"{limit_price:.2f}"
        
        for product_type in ("INTRADAY", "CARRYFORWARD"):
            try:
                order_params = {
                    "variety": "NORMAL",
                    "tradingsymbol": str(trading_symbol).strip(),
                    "symboltoken": str(symbol_token).strip(),
                    "transactiontype": "SELL",
                    "exchange": "BFO",
                    "ordertype": "LIMIT",
                    "producttype": product_type,
                    "duration": "DAY",
                    "price": limit_price_str,
                    "squareoff": "0",
                    "stoploss": "0",
                    "quantity": str(qty_val),
                }
                res = smart_api.placeOrder(order_params)
                order_id, err_msg = parse_angel_order_response(res)
                if order_id:
                    logger.info("⚡ [FAILSAFE LIMIT SELL ORDER PLACED] (%s) Price: %s | Order ID: %s", product_type, limit_price_str, order_id)
                    send_mobile_alert(f"🔴 *FAILSAFE LIMIT SELL PLACED*\nContract: *{trading_symbol}*\nQty: *{qty_val}*\nPrice: ₹{limit_price_str}\nOrder ID: `{order_id}`")
                    return order_id
                else:
                    logger.warning("⚠️ Limit sell rejected with producttype=%s: %s", product_type, err_msg)
            except Exception as exc:
                logger.warning("⚠️ Exception on limit sell with producttype=%s: %s", product_type, exc)
            
    logger.error("❌ Failsafe Sell Failed for %s %d Qty", trading_symbol, qty_val)
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
    current_epm_buffer = 0.13

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
                    entry_mfi_falling_15m = mem.get("entry_mfi_falling_15m", False)
                    active_strategy_name = mem.get("active_strategy", "Initial Dual MFI Lower Band Bounce")
                    initial_hard_sl = mem.get("initial_hard_sl", active_sl)
                    handover_history = mem.get("handover_history", [])
                    
                    saved_entry_time = mem.get("entry_time")
                    if saved_entry_time:
                        entry_time = datetime.fromisoformat(saved_entry_time)
                    
                    if bot_state == "CE_LONG":
                        active_contract = ce_contract
                    else:
                        active_contract = pe_contract
                    
                    logger.info("⚡ [RECALL ACTIVE POSITION] Resumed active %s trade from memory (Strategy: %s, Entry: ₹%.2f, SL: ₹%.2f, Target: ₹%.2f, Handovers: %d)", bot_state, active_strategy_name, active_entry_price, active_sl, active_target, len(handover_history))

            except (KeyError, ValueError, TypeError) as exc:
                logger.warning("⚠️ Saved memory schema mismatch. Resetting and calculating master grid cleanly: %s", exc)
                grid = None
                grid_from_memory = False

    if grid is None:
        logger.info("🆕 [MASTER GRID] Calculating a new Master Grid (3 ITM Strikes for CE & PE) for slot %s...", current_slot)
        grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(smart_api, current_slot, spot_price, spot_open, vix_val, buffer=current_epm_buffer)
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
            "Use these keywords to manage your bot and active trades on the go:\n\n"
            "1. *Add Lots:* `ADD LOTS`\n"
            "   Example: `ADD 2` (Adds 2 more lots at Market price)\n\n"
            "2. *Set Trailing Buffer:* `BUFFER POINTS`\n"
            "   Example: `BUFFER 10` (Sets trailing stop-loss distance to 10 points)\n\n"
            "3. *Modify Stop-Loss:* `SL PRICE`\n"
            "   Example: `SL 450` (Manually sets Stop Loss to ₹450)\n\n"
            "4. *Live / Demo Mode:* `LIVE [LOTS]` or `DEMO`\n"
            "   Example: `LIVE 2` (Switches to Live mode with 2 lots) or `DEMO`\n\n"
            "5. *Max Trades Limit:* `LIMIT [N]`\n"
            "   Example: `LIMIT 3` (Sets max trades per day limit)\n\n"
            "6. *Trade Swaps:* `Y` or `N`\n"
            "   (Accepts or declines pending trade swap signals)\n\n"
            "7. *Re-Entry Control:* `REENTRY ON` or `REENTRY OFF`\n"
            "   (Enables or disables trend continuation re-entry logic)\n\n"
            "8. *Safety Stops:* `STOP` / `HALT` / `EXIT` / `CLOSE`\n"
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

    # Continuous Monitoring Loop: 0.05s (50ms) high-frequency tick drive when WS feed is active, 1.0s HTTP fallback
    poll_interval = 0.05
    is_continuous = "--once" not in sys.argv
    execution_mode = "LIVE" if ("--live" in sys.argv or os.getenv("EXECUTION_MODE", "").upper() == "LIVE" or os.getenv("PAPER_MODE", "").lower() == "false") else "PAPER"

    # State Machine Variables
    bot_state = "IDLE"  # Options: "IDLE", "CE_LONG", "PE_LONG"
    trades_completed = 0
    max_trades_per_day = 2
    ce_sl_hit_today = False
    pe_sl_hit_today = False
    active_contract = None
    active_entry_price = 0.0
    active_sl = 0.0
    active_target = 0.0
    initial_hard_sl = 0.0
    active_strategy_name = ""
    handover_history: list[dict] = []
    entry_time = None
    entry_mfi_falling_15m = False
    lot_size = 1
    loop_counter = 0
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"

    # Instantiate Modular 3-Strategy Suite with Dynamic In-Flight Handover
    handover_strategies = [
        PreviousHighBreakoutMomentumStrategy(),
        PostSLRecoveryReentryStrategy(),
        DynamicSwingLowBreakoutRetestStrategy(),
    ]
    handover_engine = StateMachineHandoverEngine(strategies=handover_strategies, min_ev_improvement=MIN_EV_HANDOVER_DELTA)

    # Pending swap trade states
    pending_swap_signal = None  # None, "CE", or "PE"
    pending_swap_contract = None
    pending_swap_sl = 0.0
    pending_swap_target = 0.0
    pending_swap_type_str = ""
    pending_swap_time = None

    # Dynamic Swing Low & Bollinger Band State Tracking
    waiting_for_bb_pullback_ce = False
    lower_bb_touched_ce = False
    waiting_for_bb_pullback_pe = False
    lower_bb_touched_pe = False

    # Trailing Stop-Loss Variables
    allow_reentry_live = True
    initial_entry_happened = False
    recovery_reentry_eligible = False
    recovery_reentry_done = False
    base_lot_size = lot_size
    staggered_scaled_in = False
    initial_entry_price = 0.0
    original_sl_distance = 0.0
    trailing_active = False
    peak_price = 0.0
    trail_buffer = 5.0
    offloaded = False

    # Memory of lowest and highest prices observed in IDLE state
    recent_ce_low = ce_ltp
    recent_pe_low = pe_ltp
    previous_ce_high = ce_ltp
    previous_pe_high = pe_ltp
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

                # Check for slot transition (always transition and notify irrespective of positions, as requested)
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
                    grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(smart_api, current_slot, spot_price, spot_open, vix_val, buffer=current_epm_buffer)
                    ce_contract = ce_contracts[0]
                    pe_contract = pe_contracts[0]
                    ce_ltp = grid.ce_leg.ltp
                    pe_ltp = grid.pe_leg.ltp
                    dte_days = grid.dte

                    # Save EPM to memory
                    if bot_state in ("CE_LONG", "PE_LONG"):
                        save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                    else:
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)

                    # Send Telegram Notification for new slot's EPM (make sure current_slot is explicitly passed)
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

                    msg = format_grid_notification(grid, f"🔔 *SENSEX MASTER GRID UPDATED ({current_slot} Slot)*", spot_price, spot_open, vix_val, dte_days, current_slot)
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

                    # Update WebSocket Feed subscription for the new contracts and active contract
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
                            tokens_to_subscribe = ["99919000", ce_contract.symbol_token, pe_contract.symbol_token]
                            if bot_state != "IDLE" and active_contract:
                                tokens_to_subscribe.append(active_contract.symbol_token)
                            ws_feed.start(tokens_to_subscribe)
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
                    if remote_lots < 1.0:
                        current_epm_buffer = remote_lots
                        logger.info("⚙️ [TELEGRAM] EPM Master Grid Buffer updated to %.2f. Recalculating EPM Grid...", current_epm_buffer)
                        
                        grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(smart_api, current_slot, spot_price, spot_open, vix_val, buffer=current_epm_buffer)
                        ce_contract = ce_contracts[0]
                        pe_contract = pe_contracts[0]
                        ce_ltp = grid.ce_leg.ltp
                        pe_ltp = grid.pe_leg.ltp
                        dte_days = grid.dte
                        
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        msg = format_grid_notification(grid, f"🔔 *SENSEX MASTER GRID UPDATED (Buffer: {current_epm_buffer:.2f})*", spot_price, spot_open, vix_val, dte_days, current_slot)
                        send_mobile_alert(msg)
                    else:
                        trail_buffer = remote_lots
                        logger.info("⚙️ [TELEGRAM] Trailing buffer manually updated to %.1f points.", trail_buffer)
                        send_mobile_alert(f"⚙️ *TRAILING BUFFER UPDATED*\n\nBuffer manually updated to *{trail_buffer:.1f} points* via Telegram.")
                elif cmd == "LIMIT":
                    max_trades_per_day = int(remote_lots)
                    logger.info("⚙️ [TELEGRAM] Max Trades Limit manually updated to %d.", max_trades_per_day)
                    send_mobile_alert(f"⚙️ *MAX TRADES LIMIT UPDATED*\n\nMaximum trades per day manually updated to *{max_trades_per_day} trades* via Telegram.")
                elif cmd == "Y" and pending_swap_signal and bot_state in ("CE_LONG", "PE_LONG"):
                    if pending_swap_time and (datetime.now(IST) - pending_swap_time).total_seconds() <= 60.0:
                        sys.stdout.write("\n")
                        logger.info("🔄 [SWAP] User confirmed position swap to %s via Telegram!", pending_swap_signal)
                        
                        swap_ce_price = ce_contract.ltp if ce_contract else 0.0
                        swap_pe_price = pe_contract.ltp if pe_contract else 0.0
                        exit_price = swap_ce_price if bot_state == "CE_LONG" else swap_pe_price
                        logger.info("🔴 [SWAP EXIT] Exiting active %s position at ₹%.2f", active_contract.trading_symbol, exit_price)
                        send_mobile_alert(f"🔴 *SWAP EXIT: EXITING CURRENT POSITION*\n\nClosing *{active_contract.trading_symbol}* at ₹{exit_price:.2f} to switch trades.")
                        
                        if execution_mode == "LIVE":
                            execute_failsafe_sell(smart_api, active_contract.trading_symbol, active_contract.symbol_token, lot_size * 20, exit_price)
                            
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": "EXIT_SWAP",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": exit_price,
                            "qty": lot_size * 20,
                            "trades_count": trades_completed + 1
                        })
                        
                        bot_state = f"{pending_swap_signal}_LONG"
                        active_contract = pending_swap_contract
                        active_entry_price = swap_ce_price if pending_swap_signal == "CE" else swap_pe_price
                        active_target = pending_swap_target
                        active_sl = pending_swap_sl
                        entry_time = datetime.now(IST)
                        qty_to_trade = lot_size * 20
                        original_sl_distance = max(15.0, active_entry_price - active_sl)
                        trailing_active = False
                        peak_price = active_entry_price
                        offloaded = False
                        
                        logger.info("🟢 [SWAP ENTRY] Switched into %s setup at ₹%.2f (SL: ₹%.2f, Target: ₹%.2f)", bot_state, active_entry_price, active_sl, active_target)
                        send_mobile_alert(f"🟢 *SWAP ENTRY SUCCESSFUL ({pending_swap_type_str})*\n\n"
                                          f"Contract: *{active_contract.trading_symbol}*\n"
                                          f"Entry Price: ₹{active_entry_price:.2f}\n"
                                          f"Stop Loss: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}\n"
                                          f"Mode: *{execution_mode}* | Lot Size: *{lot_size}* ({qty_to_trade} Qty)")
                                          
                        if execution_mode == "LIVE":
                            submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_trade)
                            
                        excel_tracker.add_order({
                            "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode,
                            "state": "ENTRY_SWAP",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": active_entry_price,
                            "qty": qty_to_trade,
                            "trades_count": trades_completed + 1
                        })
                        if pending_swap_type_str:
                            active_strategy_name = map_entry_type_to_strategy_name(pending_swap_type_str)
                            initial_hard_sl = active_sl
                            handover_history = []
                        save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                    else:
                        send_mobile_alert("⚠️ *SWAP REQUEST EXPIRED*\nThe pending swap request has expired (60 seconds limit). Current trade maintained.")
                    
                    pending_swap_signal = None
                    pending_swap_contract = None
                    pending_swap_time = None
                elif cmd == "N" and pending_swap_signal:
                    logger.info("⚙️ [TELEGRAM] User declined trade swap.")
                    send_mobile_alert("⚙️ *SWAP REQUEST DECLINED*\nMaintaining existing position as is.")
                    pending_swap_signal = None
                    pending_swap_contract = None
                    pending_swap_time = None
                elif cmd in ("REENTRY ON", "ON REENTRY", "REENTRY TRUE"):
                    allow_reentry_live = True
                    logger.info("⚙️ [TELEGRAM] Re-entry logic enabled via Telegram.")
                    send_mobile_alert("⚙️ *RE-ENTRY LOGIC ENABLED*\n\nTrend continuation re-entries are now ACTIVE.")
                elif cmd in ("REENTRY OFF", "OFF REENTRY", "REENTRY FALSE"):
                    allow_reentry_live = False
                    logger.info("⚙️ [TELEGRAM] Re-entry logic disabled via Telegram.")
                    send_mobile_alert("⚙️ *RE-ENTRY LOGIC DISABLED*\n\nTrend continuation re-entries are now INACTIVE.")
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

                # Fetch active position LTP directly to isolate active position tracking from slot transitions
                live_active_ltp = None
                if bot_state in ("CE_LONG", "PE_LONG") and active_contract:
                    if ws_feed and ws_feed.is_connected:
                        live_active_ltp = ws_feed.prices.get(active_contract.symbol_token)
                    if live_active_ltp is None:
                        try:
                            live_active_res = smart_api.ltpData("BFO", active_contract.trading_symbol, active_contract.symbol_token)
                            live_active_ltp = float(live_active_res["data"]["ltp"]) if isinstance(live_active_res, dict) and live_active_res.get("data") else None
                        except Exception:
                            live_active_ltp = None

                if live_spot is None or live_ce_ltp is None or live_pe_ltp is None:
                    logger.warning("⚠️ [API DELAY] Live feed or LTP data returned None (likely Rate Limited). Skipping loop iteration to prevent stale trades.")
                    continue

                if bot_state in ("CE_LONG", "PE_LONG") and live_active_ltp is None:
                    logger.warning("⚠️ [API DELAY] Active trade LTP returned None. Skipping loop iteration to prevent stale trades.")
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
                        active_strategy_name = ""
                        initial_hard_sl = 0.0
                        handover_history = []
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        # Reset recent low tracking
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                # 3. State Machine Signal Evaluation (Evaluate entries when IDLE, evaluate exits when in active trade)
                if bot_state == "IDLE":
                    if trades_completed >= max_trades_per_day:
                        pass
                    else:
                        now_time = datetime.now(IST).time()
                        
                        # Technical Proximity Check (EPM removed from signal gating)
                        ce_min_dist = 0.0
                        pe_min_dist = 0.0
                        ce_in_15_range = True
                        pe_in_15_range = True
                        
                        weekly_open = get_weekly_open_price(smart_api)
                        
                        # Priority check based on price action and momentum
                        check_order = ["CE", "PE"] if (live_spot >= weekly_open) else ["PE", "CE"]
                        
                        ce_entry_signal = False
                        active_sl_ce = live_ce_ltp - 20.0
                        entry_type_str_ce = ""
                        
                        pe_entry_signal = False
                        active_sl_pe = live_pe_ltp - 20.0
                        entry_type_str_pe = ""
                        
                        now_time_str = datetime.now(IST).strftime("%H:%M")
                        is_915_opening = (now_time_str <= "09:20")

                        # Rule: Restrict fresh Entry on or after 15:00:00
                        if now_time_str >= "15:00":
                            if loop_counter % 20 == 1:
                                logger.info("⏸️ [15:00 ENTRY CUTOFF] Fresh entries blocked on or after 15:00:00 (Current Time: %s). Monitoring existing trades.", now_time_str)
                            time.sleep(5)
                            continue
                        
                        for opt_type in check_order:
                            if opt_type == "CE" and not ce_entry_signal and not pe_entry_signal:
                                if ce_sl_hit_today:
                                    logger.info("⛔ [CE SIDE LOCKED] CE Stop Loss was hit today. Skipping fresh CE entries.")
                                    continue
                                c_open_15m, c_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", ce_contract.symbol_token)
                                if c_open_15m is None:
                                    c_open_15m = live_ce_ltp
                                if c_low_15m is None:
                                    c_low_15m = min(recent_ce_low, live_ce_ltp)
                                else:
                                    c_low_15m = min(c_low_15m, recent_ce_low)
                                    
                                mfis_15m, prev_mfis_15m = get_mfi_multi_period(smart_api, "BFO", ce_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                                mfis_3m, prev_mfis_3m = get_mfi_multi_period(smart_api, "BFO", ce_contract.symbol_token, "THREE_MINUTE", [5, 14])
                                mfis_1m, prev_mfis_1m = get_mfi_multi_period(smart_api, "BFO", ce_contract.symbol_token, "ONE_MINUTE", [5, 14])
                                mfis_30m, prev_mfis_30m, prev_prev_mfis_30m, extra_30m_ce = get_mfi_multi_period(smart_api, "BFO", ce_contract.symbol_token, "THIRTY_MINUTE", [5, 14], return_extra=True)
                                mfis_60m, prev_mfis_60m = get_mfi_multi_period(smart_api, "BFO", ce_contract.symbol_token, "ONE_HOUR", [5, 14])
                                mb_3m_ce, ub_3m_ce, lb_3m_ce, c_open_3m_ce, c_low_3m_ce, c_close_3m_ce, prev_c_close_3m_ce, swing_low_3m_ce = get_3m_bollinger_bands(smart_api, "BFO", ce_contract.symbol_token)
                                
                                mfi5_15m = mfis_15m.get(5, 50.0)
                                mfi14_15m = mfis_15m.get(14, 50.0)
                                prev_mfi5_15m = prev_mfis_15m.get(5, 50.0)
                                prev_mfi14_15m = prev_mfis_15m.get(14, 50.0)

                                mfi5_3m = mfis_3m.get(5, 50.0)
                                mfi14_3m = mfis_3m.get(14, 50.0)
                                prev_mfi5_3m = prev_mfis_3m.get(5, 50.0)
                                prev_mfi14_3m = prev_mfis_3m.get(14, 50.0)
                                
                                mfi5_1m = mfis_1m.get(5, 50.0)
                                mfi14_1m = mfis_1m.get(14, 50.0)
                                prev_mfi5_1m = prev_mfis_1m.get(5, 50.0)
                                prev_mfi14_1m = prev_mfis_1m.get(14, 50.0)

                                mfi5_30m = mfis_30m.get(5, 50.0)
                                mfi14_30m = mfis_30m.get(14, 50.0)
                                prev_mfi5_30m = prev_mfis_30m.get(5, 50.0)
                                prev_mfi14_30m = prev_mfis_30m.get(14, 50.0)
                                prev_prev_mfi5_30m = prev_prev_mfis_30m.get(5, 50.0)
                                prev_prev_mfi14_30m = prev_prev_mfis_30m.get(14, 50.0)
                                
                                prev_close_30m_ce = extra_30m_ce.get("prev_close", 0.0)
                                prev_prev_low_30m_ce = extra_30m_ce.get("prev_prev_low", 0.0)

                                mfi5_60m = mfis_60m.get(5, 50.0)
                                mfi14_60m = mfis_60m.get(14, 50.0)
                                prev_mfi5_60m = prev_mfis_60m.get(5, 50.0)
                                prev_mfi14_60m = prev_mfis_60m.get(14, 50.0)

                                # Higher Timeframe Block Filters:
                                # 1. Hourly chart both MFI falling (60m MFI5 & MFI14 falling)
                                is_60m_both_falling_ce = (mfi5_60m < prev_mfi5_60m) and (mfi14_60m < prev_mfi14_60m)
                                # 2. 30m MFI 5 falling after reaching 100/overbought (>= 90)
                                is_30m_overbought_falling_ce = (mfi5_30m >= 90.0 or prev_mfi5_30m >= 90.0) and (mfi5_30m < prev_mfi5_30m)
                                higher_tf_block_ce = is_60m_both_falling_ce or is_30m_overbought_falling_ce
                                
                                # 15m MFI Pattern for Initial Entry:
                                # Previous MFI(5) was strictly 0.0 and currently bouncing (> 0.0) + MFI(14) <= 25.0 rising
                                is_mfi5_zero_bounce_ce = (prev_mfi5_15m == 0.0 and mfi5_15m > 0.0) or (mfi5_15m == 0.0)
                                is_mfi14_bounce_ce = (mfi14_15m <= 25.0) and (mfi14_15m >= prev_mfi14_15m)
                                is_initial_mfi_pattern_ce = is_mfi5_zero_bounce_ce and is_mfi14_bounce_ce
                                
                                # Reverse Oversold Exception: If Spot < Weekly Open but CE is extremely oversold in 15m and 30m
                                is_dual_oversold_ce = (mfi5_30m == 0.0 and mfi14_30m <= 25.0)
                                is_direction_aligned_ce = (live_spot >= weekly_open) or is_dual_oversold_ce

                                # High Alert notification when MFI(5) is strictly 0.0 on 15m candle
                                if mfi5_15m == 0.0 and loop_counter % 30 == 1:
                                    send_mobile_alert(
                                        f"🚨 *HIGH ALERT: 15-MIN MFI(5) AT ZERO*\n\n"
                                        f"Contract: *{ce_contract.trading_symbol}*\n"
                                        f"15m MFI(5): *{mfi5_15m:.1f}* | MFI(14): *{mfi14_15m:.1f}*\n"
                                        f"LTP: ₹{live_ce_ltp:.2f}\n"
                                        f"👉 *Action*: Stay on High Alert! Monitoring for MFI(5) zero bounce entry."
                                    )

                                is_bounce_open_ce = (live_ce_ltp >= c_open_15m)
                                
                                # 1. Dual 15m MFI Pattern for Initial Entry (Must be below Middle Band near Lower Band & 30m MFI 14 not falling)
                                is_below_mb_ce = (live_ce_ltp < c_open_15m or live_ce_ltp <= previous_ce_high)
                                is_30m_mfi14_not_falling_ce = (mfi14_30m >= prev_mfi14_30m)
                                is_clean_initial_entry_ce = is_initial_mfi_pattern_ce and is_below_mb_ce and is_30m_mfi14_not_falling_ce

                                # 2. Previous High Breakout Momentum Entry Option:
                                is_30m_both_favorable_ce = (mfi5_30m >= prev_mfi5_30m) and (mfi14_30m >= prev_mfi14_30m + 0.5)
                                is_both_15m_mfi_rising_ce = (mfi5_15m > prev_mfi5_15m) and (mfi14_15m >= prev_mfi14_15m)
                                if is_below_mb_ce:
                                    is_breakout_entry_ce = is_both_15m_mfi_rising_ce and is_bounce_open_ce and (mfi14_15m > 25.0) and (live_ce_ltp > previous_ce_high) and is_30m_both_favorable_ce
                                else:
                                    is_breakout_entry_ce = is_both_15m_mfi_rising_ce and (mfi14_15m > 25.0) and (live_ce_ltp > previous_ce_high) and is_30m_both_favorable_ce
                                
                                # Rule #3 Entry Filter for Breakout:
                                is_breakout_opening_too_high_ce = (live_ce_ltp - (c_low_15m - 4.0)) > 20.0 or (live_ce_ltp - previous_ce_high) >= 50.0
                                is_mfi_overbought_100_ce = (mfi5_15m >= 100.0 or mfi14_15m >= 100.0)
                                if is_breakout_opening_too_high_ce or is_mfi_overbought_100_ce:
                                    is_3m_mfi_rising_corr_ce = (mfi5_1m > prev_mfi5_1m + 1.0) and (mfi14_1m >= prev_mfi14_1m)
                                    is_breakout_entry_ce = is_breakout_entry_ce and is_3m_mfi_rising_corr_ce

                                # 3. 30-min Dual MFI Reversal Option: Secondary entry only after initial setup, MFI 14 MUST be rising
                                is_last_30m_breakdown_ce = (prev_close_30m_ce < prev_prev_low_30m_ce) or (c_open_15m > live_ce_ltp)
                                is_15m_both_increasing_ce = (mfi5_15m > prev_mfi5_15m) and (mfi14_15m > prev_mfi14_15m)
                                is_30m_mfi_option_ce = initial_entry_happened and is_15m_both_increasing_ce and (mfi14_15m <= 35.0) and is_bounce_open_ce and not is_last_30m_breakdown_ce

                                # 4. Re-Entry Condition ("MFI(14) Trend Re-Entry (MB Consolidation)"):
                                is_no_mfi_falling_htf_ce = (mfi5_30m >= prev_mfi5_30m and mfi14_30m >= prev_mfi14_30m and mfi5_60m >= prev_mfi5_60m and mfi14_60m >= prev_mfi14_60m)
                                
                                is_30m_both_falling_ce = (mfi5_30m < prev_mfi5_30m) and (mfi14_30m < prev_mfi14_30m)
                                is_15m_any_mfi_ob_ce = (mfi5_15m >= 80.0 or mfi14_15m >= 68.0)
                                is_30m_mfi14_falling_prev_or_curr_ce = (mfi14_30m < prev_mfi14_30m) or (prev_mfi14_30m < prev_prev_mfi14_30m)
                                is_reentry_blocked_by_30m_ce = (
                                    is_15m_any_mfi_ob_ce or
                                    is_30m_mfi14_falling_prev_or_curr_ce or
                                    (mfi5_30m == 100.0 or mfi14_30m >= 70.0) or
                                    is_30m_both_falling_ce or
                                    is_last_30m_breakdown_ce
                                )
                                is_reentry_ce = allow_reentry_live and is_bounce_open_ce and (mfi5_15m > prev_mfi5_15m and mfi14_15m >= prev_mfi14_15m) and is_no_mfi_falling_htf_ce and not is_reentry_blocked_by_30m_ce

                                # 5. One-Time Post-SL Recovery Re-Entry (+2 Lots) with Lower Band proximity and 30m both MFIs favorable
                                is_recovery_reentry_ce = recovery_reentry_eligible and not recovery_reentry_done and is_bounce_open_ce and (mfi5_15m > prev_mfi5_15m and mfi14_15m > prev_mfi14_15m) and is_below_mb_ce and is_30m_both_favorable_ce

                                # 6. Post-Breakdown Oversold Bounce Entry:
                                is_prev_breakdown_candle_ce = (c_open_15m > live_ce_ltp) or (prev_close_30m_ce < prev_prev_low_30m_ce)
                                is_oversold_15m_mfi_ce = (mfi5_15m <= 25.0 or mfi14_15m <= 30.0)
                                is_any_mfi_increasing_15m_ce = (mfi5_15m > prev_mfi5_15m or mfi14_15m > prev_mfi14_15m)
                                is_1m_mfi_bounce_ce = (mfi5_1m > prev_mfi5_1m and mfi14_1m >= prev_mfi14_1m)
                                is_15m_both_falling_ce = (mfi5_15m < prev_mfi5_15m and mfi14_15m < prev_mfi14_15m)
                                is_post_breakdown_entry_ce = is_prev_breakdown_candle_ce and is_oversold_15m_mfi_ce and is_any_mfi_increasing_15m_ce and is_1m_mfi_bounce_ce and not is_15m_both_falling_ce and not is_30m_both_falling_ce

                                # 7. Dynamic Agent-based Swing Low First Breakout Retest Entry:
                                dynamic_swing_low_ce = min(c_low_15m, recent_ce_low, swing_low_3m_ce if swing_low_3m_ce is not None else c_low_15m)
                                dynamic_range_ce = max(15.0, previous_ce_high - dynamic_swing_low_ce)
                                dynamic_tolerance_ce = max(3.0, min(12.0, 0.15 * dynamic_range_ce))

                                is_mfi_falling_from_ob_ce = (prev_mfi5_15m >= 80.0 or prev_mfi14_15m >= 70.0) and (mfi5_15m < prev_mfi5_15m or mfi14_15m < prev_mfi14_15m)
                                is_swing_ob_blocked_ce = (mfi5_15m >= 80.0 or mfi14_15m >= 68.0) or (ub_3m_ce is not None and live_ce_ltp >= ub_3m_ce - 8.0) or is_mfi_falling_from_ob_ce

                                if is_swing_ob_blocked_ce:
                                    waiting_for_bb_pullback_ce = False
                                    lower_bb_touched_ce = False

                                # 3m MFI condition: Both MFI increasing OR at least MFI(14) increasing
                                is_3m_both_falling_ce = (mfi5_3m < prev_mfi5_3m) and (mfi14_3m < prev_mfi14_3m)
                                is_3m_mfi_rising_ce = ((mfi5_3m > prev_mfi5_3m) and (mfi14_3m >= prev_mfi14_3m)) or (mfi14_3m > prev_mfi14_3m)
                                
                                # HTF MFI Rising check: 15m or 30m MFI rising
                                is_15m_mfi_rising_ce = (mfi14_15m > prev_mfi14_15m) or (mfi14_15m >= prev_mfi14_15m and mfi5_15m > prev_mfi5_15m)
                                is_30m_mfi_rising_ce = (mfi14_30m > prev_mfi14_30m) or (mfi14_30m >= prev_mfi14_30m and mfi5_30m > prev_mfi5_30m)
                                is_htf_mfi_rising_ce = is_15m_mfi_rising_ce or is_30m_mfi_rising_ce

                                # Track Lower Bollinger Band touch
                                if lb_3m_ce is not None and (live_ce_ltp <= lb_3m_ce + 2.0 or (c_low_3m_ce is not None and c_low_3m_ce <= lb_3m_ce + 2.0) or c_low_15m <= lb_3m_ce + 2.0):
                                    lower_bb_touched_ce = True

                                # Middle Band Proximity & Correction at or below MB:
                                is_corrected_to_mb_ce = False
                                if mb_3m_ce is not None:
                                    is_corrected_to_mb_ce = (live_ce_ltp <= mb_3m_ce + 1.0 or (c_low_3m_ce is not None and c_low_3m_ce <= mb_3m_ce + 1.0) or c_low_15m <= mb_3m_ce + 1.0)
                                
                                # Middle Band entry condition: Correct at/below MB, bounce to open, both 3m MFI (or MFI14) rising AND any HTF MFI rising
                                is_mb_swing_entry_ce = is_corrected_to_mb_ce and is_bounce_open_ce and is_3m_mfi_rising_ce and is_htf_mfi_rising_ce and not is_swing_ob_blocked_ce

                                # If 3m MFIs are falling, force waiting for lower Bollinger Band touch
                                if is_3m_both_falling_ce:
                                    waiting_for_bb_pullback_ce = True

                                is_lower_bb_bounce_entry_ce = waiting_for_bb_pullback_ce and lower_bb_touched_ce and is_bounce_open_ce and is_3m_mfi_rising_ce and is_htf_mfi_rising_ce and not is_swing_ob_blocked_ce

                                # Retest near validated dynamic swing low
                                is_near_swing_low_ce = (c_low_15m <= dynamic_swing_low_ce + (2.5 * dynamic_tolerance_ce)) or (live_ce_ltp <= dynamic_swing_low_ce + (2.5 * dynamic_tolerance_ce))
                                is_direct_swing_entry_ce = not waiting_for_bb_pullback_ce and is_near_swing_low_ce and is_bounce_open_ce and is_3m_mfi_rising_ce and is_htf_mfi_rising_ce and not is_swing_ob_blocked_ce

                                is_swing_low_retest_entry_ce = (is_mb_swing_entry_ce or is_lower_bb_bounce_entry_ce or is_direct_swing_entry_ce) and not is_30m_both_falling_ce and not is_swing_ob_blocked_ce

                                if is_swing_low_retest_entry_ce:
                                    waiting_for_bb_pullback_ce = False
                                    lower_bb_touched_ce = False

                                # Rule #1 Entry Filter: Block fresh entry if 15m MFI(5) is at 100 / extreme overbought (>=99.0) OR MFI(14) >= 70.0 OR MFI(14) not increasing when MFI(5) >= 90
                                is_universal_ob_blocked_ce = (mfi5_15m >= 99.0) or (mfi14_15m >= 70.0) or (mfi5_15m >= 90.0 and not (mfi14_15m > prev_mfi14_15m))
                                if is_universal_ob_blocked_ce:
                                    is_clean_initial_entry_ce = False
                                    is_breakout_entry_ce = False
                                    is_30m_mfi_option_ce = False
                                    is_reentry_ce = False
                                    is_recovery_reentry_ce = False
                                    is_post_breakdown_entry_ce = False
                                    is_swing_low_retest_entry_ce = False

                                # 9:15 Big Gap Up Retest & Bounce requirement (Applies universally to ALL entry types)
                                is_ce_big_gap_up = (c_open_15m - previous_ce_high >= 30.0) or (live_ce_ltp - previous_ce_high >= 30.0)
                                if is_915_opening and is_ce_big_gap_up:
                                    is_clean_initial_entry_ce = False
                                    is_breakout_entry_ce = False
                                    is_30m_mfi_option_ce = False
                                    is_reentry_ce = False
                                    is_recovery_reentry_ce = False
                                    is_post_breakdown_entry_ce = False
                                    is_swing_low_retest_entry_ce = False
                                elif is_ce_big_gap_up and not is_915_opening:
                                    is_near_low_35pt_ce = (live_ce_ltp <= c_low_15m + 35.0)
                                    is_mfi_rising_3m_ce = (mfi5_1m > prev_mfi5_1m) and (mfi14_1m >= prev_mfi14_1m)
                                    is_extreme_oversold_reversal_ce = (mfi5_15m == 0.0 and mfi14_15m <= 15.0) and (live_ce_ltp >= c_low_15m + 7.0)
                                    is_bounce_to_open_ce = (live_ce_ltp >= c_open_15m) and (mfi5_1m > prev_mfi5_1m and mfi14_1m > prev_mfi14_1m)
                                    
                                    is_valid_retest_ce = is_near_low_35pt_ce and (is_mfi_rising_3m_ce or is_extreme_oversold_reversal_ce or is_bounce_to_open_ce)
                                    if not is_valid_retest_ce:
                                        is_clean_initial_entry_ce = False
                                        is_breakout_entry_ce = False
                                        is_30m_mfi_option_ce = False
                                        is_reentry_ce = False
                                        is_recovery_reentry_ce = False
                                        is_post_breakdown_entry_ce = False
                                        is_swing_low_retest_entry_ce = False

                                # --- Dynamic EPM Low Bounce LONG Entry Signal (CE) ---
                                ce_epm_low_saved = load_grid_state_epm_low("CE")
                                if ce_epm_low_saved is None or ce_epm_low_saved <= 0:
                                    ce_epm_low_saved = grid.ce_leg.epm_lower_range if (grid and grid.ce_leg) else 0.0

                                atr14_ce, stddev20_ce = calculate_atr_and_stddev(smart_api, "BFO", ce_contract.symbol_token)
                                dynamic_near_thresh_ce = max(0.25 * atr14_ce, 0.8 * stddev20_ce, 12.0)

                                is_price_near_epm_low_ce = (ce_epm_low_saved > 0.0) and (live_ce_ltp >= ce_epm_low_saved) and ((live_ce_ltp - ce_epm_low_saved) <= dynamic_near_thresh_ce)
                                prev_low_ce_check = c_low_15m if c_low_15m is not None else live_ce_ltp
                                is_bouncing_ce = (ce_epm_low_saved > 0.0) and (prev_low_ce_check <= ce_epm_low_saved + dynamic_near_thresh_ce) and (live_ce_ltp >= c_open_15m)
                                is_mfi_increasing_ce = (mfi14_15m > prev_mfi14_15m) or (mfi5_15m == 0.0) or (prev_mfi5_15m == 0.0 and mfi5_15m > 0.0) or (mfi14_15m <= 25.0 and mfi14_15m > prev_mfi14_15m)

                                is_epm_low_bounce_entry_ce = is_price_near_epm_low_ce and is_bouncing_ce and is_mfi_increasing_ce

                                if not higher_tf_block_ce:
                                    if is_epm_low_bounce_entry_ce:
                                        ce_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_ce = live_ce_ltp - 20.0
                                        entry_type_str_ce = f"CE Dynamic EPM Low Bounce LONG Entry (EPM Low: ₹{ce_epm_low_saved:.2f}, Thresh: ₹{dynamic_near_thresh_ce:.1f} | SL-20)"
                                    elif is_recovery_reentry_ce:
                                        ce_entry_signal = True
                                        recovery_reentry_eligible = False
                                        recovery_reentry_done = True
                                        lot_size = base_lot_size + 2
                                        active_sl_ce = live_ce_ltp - 20.0
                                        entry_type_str_ce = f"CE One-Time Post-SL Recovery Re-Entry (+2 Lots, Total: {lot_size} Lots | SL-20)"
                                    elif is_swing_low_retest_entry_ce:
                                        ce_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_ce = max(dynamic_swing_low_ce - 2.0, live_ce_ltp - 20.0)
                                        entry_type_str_ce = f"CE Dynamic Swing Low First Breakout Retest Entry (Pivot: ₹{dynamic_swing_low_ce:.2f}, MFI14={mfi14_15m:.1f} | SL: ₹{active_sl_ce:.2f})"
                                    elif is_breakout_entry_ce:
                                        ce_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        if (mfi5_15m > prev_mfi5_15m and mfi14_15m > prev_mfi14_15m):
                                            active_sl_ce = max(c_low_15m - 15.0, live_ce_ltp - 20.0)
                                        else:
                                            active_sl_ce = live_ce_ltp - 20.0
                                        entry_type_str_ce = f"CE Previous High Breakout Momentum Entry (MFI14={mfi14_15m:.1f} | SL: ₹{active_sl_ce:.2f})"
                                    elif is_direction_aligned_ce and is_clean_initial_entry_ce:
                                        ce_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_ce = live_ce_ltp - 20.0
                                        entry_type_str_ce = f"CE Initial MFI Bounce (5=0 & 14<={mfi14_15m:.1f} | SL-20 from Entry)"

                            elif opt_type == "PE" and not ce_entry_signal and not pe_entry_signal:
                                if pe_sl_hit_today:
                                    logger.info("⛔ [PE SIDE LOCKED] PE Stop Loss was hit today. Skipping fresh PE entries.")
                                    continue
                                p_open_15m, p_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", pe_contract.symbol_token)
                                if p_open_15m is None:
                                    p_open_15m = live_pe_ltp
                                if p_low_15m is None:
                                    p_low_15m = min(recent_pe_low, live_pe_ltp)
                                else:
                                    p_low_15m = min(p_low_15m, recent_pe_low)
                                    
                                mfis_15m_pe, prev_mfis_15m_pe = get_mfi_multi_period(smart_api, "BFO", pe_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                                mfis_3m_pe, prev_mfis_3m_pe = get_mfi_multi_period(smart_api, "BFO", pe_contract.symbol_token, "THREE_MINUTE", [5, 14])
                                mfis_1m_pe, prev_mfis_1m_pe = get_mfi_multi_period(smart_api, "BFO", pe_contract.symbol_token, "ONE_MINUTE", [5, 14])
                                mfis_30m_pe, prev_mfis_30m_pe, prev_prev_mfis_30m_pe, extra_30m_pe = get_mfi_multi_period(smart_api, "BFO", pe_contract.symbol_token, "THIRTY_MINUTE", [5, 14], return_extra=True)
                                mfis_60m_pe, prev_mfis_60m_pe = get_mfi_multi_period(smart_api, "BFO", pe_contract.symbol_token, "ONE_HOUR", [5, 14])
                                mb_3m_pe, ub_3m_pe, lb_3m_pe, c_open_3m_pe, c_low_3m_pe, c_close_3m_pe, prev_c_close_3m_pe, swing_low_3m_pe = get_3m_bollinger_bands(smart_api, "BFO", pe_contract.symbol_token)
                                
                                mfi5_15m_pe = mfis_15m_pe.get(5, 50.0)
                                mfi14_15m_pe = mfis_15m_pe.get(14, 50.0)
                                prev_mfi5_15m_pe = prev_mfis_15m_pe.get(5, 50.0)
                                prev_mfi14_15m_pe = prev_mfis_15m_pe.get(14, 50.0)

                                mfi5_3m_pe = mfis_3m_pe.get(5, 50.0)
                                mfi14_3m_pe = mfis_3m_pe.get(14, 50.0)
                                prev_mfi5_3m_pe = prev_mfis_3m_pe.get(5, 50.0)
                                prev_mfi14_3m_pe = prev_mfis_3m_pe.get(14, 50.0)
                                
                                mfi5_1m_pe = mfis_1m_pe.get(5, 50.0)
                                mfi14_1m_pe = mfis_1m_pe.get(14, 50.0)
                                prev_mfi5_1m_pe = prev_mfis_1m_pe.get(5, 50.0)
                                prev_mfi14_1m_pe = prev_mfis_1m_pe.get(14, 50.0)

                                mfi5_30m_pe = mfis_30m_pe.get(5, 50.0)
                                mfi14_30m_pe = mfis_30m_pe.get(14, 50.0)
                                prev_mfi5_30m_pe = prev_mfis_30m_pe.get(5, 50.0)
                                prev_mfi14_30m_pe = prev_mfis_30m_pe.get(14, 50.0)
                                prev_prev_mfi5_30m_pe = prev_prev_mfis_30m_pe.get(5, 50.0)
                                prev_prev_mfi14_30m_pe = prev_prev_mfis_30m_pe.get(14, 50.0)
                                
                                prev_close_30m_pe = extra_30m_pe.get("prev_close", 0.0)
                                prev_prev_low_30m_pe = extra_30m_pe.get("prev_prev_low", 0.0)

                                mfi5_60m_pe = mfis_60m_pe.get(5, 50.0)
                                mfi14_60m_pe = mfis_60m_pe.get(14, 50.0)
                                prev_mfi5_60m_pe = prev_mfis_60m_pe.get(5, 50.0)
                                prev_mfi14_60m_pe = prev_mfis_60m_pe.get(14, 50.0)

                                # Higher Timeframe Block Filters for PE:
                                # 1. Hourly chart both MFI falling (60m MFI5 & MFI14 falling)
                                is_60m_both_falling_pe = (mfi5_60m_pe < prev_mfi5_60m_pe) and (mfi14_60m_pe < prev_mfi14_60m_pe)
                                # 2. 30m MFI 5 falling after reaching 100/overbought (>= 90)
                                is_30m_overbought_falling_pe = (mfi5_30m_pe >= 90.0 or prev_mfi5_30m_pe >= 90.0) and (mfi5_30m_pe < prev_mfi5_30m_pe)
                                higher_tf_block_pe = is_60m_both_falling_pe or is_30m_overbought_falling_pe

                                # 15m MFI Pattern for Initial Entry:
                                # Previous MFI(5) was strictly 0.0 and currently bouncing (> 0.0) + MFI(14) <= 25.0 rising
                                is_mfi5_zero_bounce_pe = (prev_mfi5_15m_pe == 0.0 and mfi5_15m_pe > 0.0) or (mfi5_15m_pe == 0.0)
                                is_mfi14_bounce_pe = (mfi14_15m_pe <= 25.0) and (mfi14_15m_pe >= prev_mfi14_15m_pe)
                                is_initial_mfi_pattern_pe = is_mfi5_zero_bounce_pe and is_mfi14_bounce_pe

                                # Reverse Oversold Exception: If Spot > Weekly Open but PE is extremely oversold in 15m and 30m
                                is_dual_oversold_pe = (mfi5_30m_pe == 0.0 and mfi14_30m_pe <= 25.0)
                                is_direction_aligned_pe = (live_spot <= weekly_open) or is_dual_oversold_pe
                                
                                # High Alert notification when MFI(5) is strictly 0.0 on 15m candle
                                if mfi5_15m_pe == 0.0 and loop_counter % 30 == 1:
                                    send_mobile_alert(
                                        f"🚨 *HIGH ALERT: 15-MIN MFI(5) AT ZERO*\n\n"
                                        f"Contract: *{pe_contract.trading_symbol}*\n"
                                        f"15m MFI(5): *{mfi5_15m_pe:.1f}* | MFI(14): *{mfi14_15m_pe:.1f}*\n"
                                        f"LTP: ₹{live_pe_ltp:.2f}\n"
                                        f"👉 *Action*: Stay on High Alert! Monitoring for MFI(5) zero bounce entry."
                                    )

                                is_bounce_open_pe = (live_pe_ltp >= p_open_15m)

                                # 1. Dual 15m MFI Pattern for Initial Entry (Must be below Middle Band near Lower Band & 30m MFI 14 not falling)
                                is_below_mb_pe = (live_pe_ltp < p_open_15m or live_pe_ltp <= previous_pe_high)
                                is_30m_mfi14_not_falling_pe = (mfi14_30m_pe >= prev_mfi14_30m_pe)
                                is_clean_initial_entry_pe = is_initial_mfi_pattern_pe and is_below_mb_pe and is_30m_mfi14_not_falling_pe

                                # 2. Previous High Breakout Momentum Entry Option:
                                is_30m_both_favorable_pe = (mfi5_30m_pe >= prev_mfi5_30m_pe) and (mfi14_30m_pe >= prev_mfi14_30m_pe + 0.5)
                                is_both_15m_mfi_rising_pe = (mfi5_15m_pe > prev_mfi5_15m_pe) and (mfi14_15m_pe >= prev_mfi14_15m_pe)
                                if is_below_mb_pe:
                                    is_breakout_entry_pe = is_both_15m_mfi_rising_pe and is_bounce_open_pe and (mfi14_15m_pe > 25.0) and (live_pe_ltp > previous_pe_high) and is_30m_both_favorable_pe
                                else:
                                    is_breakout_entry_pe = is_both_15m_mfi_rising_pe and (mfi14_15m_pe > 25.0) and (live_pe_ltp > previous_pe_high) and is_30m_both_favorable_pe
                                
                                # Rule #3 Entry Filter for Breakout PE:
                                is_breakout_opening_too_high_pe = (live_pe_ltp - (p_low_15m - 4.0)) > 20.0 or (live_pe_ltp - previous_pe_high) >= 50.0
                                is_mfi_overbought_100_pe = (mfi5_15m_pe >= 100.0 or mfi14_15m_pe >= 100.0)
                                if is_breakout_opening_too_high_pe or is_mfi_overbought_100_pe:
                                    is_3m_mfi_rising_corr_pe = (mfi5_1m_pe > prev_mfi5_1m_pe + 1.0) and (mfi14_1m_pe >= prev_mfi14_1m_pe)
                                    is_breakout_entry_pe = is_breakout_entry_pe and is_3m_mfi_rising_corr_pe

                                # 3. 30-min Dual MFI Reversal Option: Secondary entry only after initial setup, MFI 14 MUST be rising
                                is_last_30m_breakdown_pe = (prev_close_30m_pe < prev_prev_low_30m_pe) or (p_open_15m > live_pe_ltp)
                                is_15m_both_increasing_pe = (mfi5_15m_pe > prev_mfi5_15m_pe) and (mfi14_15m_pe > prev_mfi14_15m_pe)
                                is_30m_mfi_option_pe = initial_entry_happened and is_15m_both_increasing_pe and (mfi14_15m_pe <= 35.0) and is_bounce_open_pe and not is_last_30m_breakdown_pe
                                
                                # 4. Re-Entry Condition ("MFI(14) Trend Re-Entry (MB Consolidation)"):
                                is_no_mfi_falling_htf_pe = (mfi5_30m_pe >= prev_mfi5_30m_pe and mfi14_30m_pe >= prev_mfi14_30m_pe and mfi5_60m_pe >= prev_mfi5_60m_pe and mfi14_60m_pe >= prev_mfi14_60m_pe)
                                
                                is_30m_both_falling_pe = (mfi5_30m_pe < prev_mfi5_30m_pe) and (mfi14_30m_pe < prev_mfi14_30m_pe)
                                is_15m_any_mfi_ob_pe = (mfi5_15m_pe >= 80.0 or mfi14_15m_pe >= 68.0)
                                is_30m_mfi14_falling_prev_or_curr_pe = (mfi14_30m_pe < prev_mfi14_30m_pe) or (prev_mfi14_30m_pe < prev_prev_mfi14_30m_pe)
                                is_reentry_blocked_by_30m_pe = (
                                    is_15m_any_mfi_ob_pe or
                                    is_30m_mfi14_falling_prev_or_curr_pe or
                                    (mfi5_30m_pe == 100.0 or mfi14_30m_pe >= 70.0) or
                                    is_30m_both_falling_pe or
                                    is_last_30m_breakdown_pe
                                )
                                is_reentry_pe = allow_reentry_live and is_bounce_open_pe and (mfi5_15m_pe > prev_mfi5_15m_pe and mfi14_15m_pe >= prev_mfi14_15m_pe) and is_no_mfi_falling_htf_pe and not is_reentry_blocked_by_30m_pe

                                # 5. One-Time Post-SL Recovery Re-Entry (+2 Lots) with Lower Band proximity and 30m both MFIs favorable
                                is_recovery_reentry_pe = recovery_reentry_eligible and not recovery_reentry_done and is_bounce_open_pe and (mfi5_15m_pe > prev_mfi5_15m_pe and mfi14_15m_pe > prev_mfi14_15m_pe) and is_below_mb_pe and is_30m_both_favorable_pe

                                # 6. Post-Breakdown Oversold Bounce Entry:
                                is_prev_breakdown_candle_pe = (p_open_15m > live_pe_ltp) or (prev_close_30m_pe < prev_prev_low_30m_pe)
                                is_oversold_15m_mfi_pe = (mfi5_15m_pe <= 25.0 or mfi14_15m_pe <= 30.0)
                                is_any_mfi_increasing_15m_pe = (mfi5_15m_pe > prev_mfi5_15m_pe or mfi14_15m_pe > prev_mfi14_15m_pe)
                                is_1m_mfi_bounce_pe = (mfi5_1m_pe > prev_mfi5_1m_pe and mfi14_1m_pe >= prev_mfi14_1m_pe)
                                is_15m_both_falling_pe = (mfi5_15m_pe < prev_mfi5_15m_pe and mfi14_15m_pe < prev_mfi14_15m_pe)
                                is_post_breakdown_entry_pe = is_prev_breakdown_candle_pe and is_oversold_15m_mfi_pe and is_any_mfi_increasing_15m_pe and is_1m_mfi_bounce_pe and not is_15m_both_falling_pe and not is_30m_both_falling_pe

                                # 7. Dynamic Agent-based Swing Low First Breakout Retest Entry for PE:
                                dynamic_swing_low_pe = min(p_low_15m, recent_pe_low, swing_low_3m_pe if swing_low_3m_pe is not None else p_low_15m)
                                dynamic_range_pe = max(15.0, previous_pe_high - dynamic_swing_low_pe)
                                dynamic_tolerance_pe = max(3.0, min(12.0, 0.15 * dynamic_range_pe))

                                is_mfi_falling_from_ob_pe = (prev_mfi5_15m_pe >= 80.0 or prev_mfi14_15m_pe >= 70.0) and (mfi5_15m_pe < prev_mfi5_15m_pe or mfi14_15m_pe < prev_mfi14_15m_pe)
                                is_swing_ob_blocked_pe = (mfi5_15m_pe >= 80.0 or mfi14_15m_pe >= 68.0) or (ub_3m_pe is not None and live_pe_ltp >= ub_3m_pe - 8.0) or is_mfi_falling_from_ob_pe

                                if is_swing_ob_blocked_pe:
                                    waiting_for_bb_pullback_pe = False
                                    lower_bb_touched_pe = False

                                # 3m MFI condition: Both MFI increasing OR at least MFI(14) increasing
                                is_3m_both_falling_pe = (mfi5_3m_pe < prev_mfi5_3m_pe) and (mfi14_3m_pe < prev_mfi14_3m_pe)
                                is_3m_mfi_rising_pe = ((mfi5_3m_pe > prev_mfi5_3m_pe) and (mfi14_3m_pe >= prev_mfi14_3m_pe)) or (mfi14_3m_pe > prev_mfi14_3m_pe)
                                
                                # HTF MFI Rising check: 15m or 30m MFI rising
                                is_15m_mfi_rising_pe = (mfi14_15m_pe > prev_mfi14_15m_pe) or (mfi14_15m_pe >= prev_mfi14_15m_pe and mfi5_15m_pe > prev_mfi5_15m_pe)
                                is_30m_mfi_rising_pe = (mfi14_30m_pe > prev_mfi14_30m_pe) or (mfi14_30m_pe >= prev_mfi14_30m_pe and mfi5_30m_pe > prev_mfi5_30m_pe)
                                is_htf_mfi_rising_pe = is_15m_mfi_rising_pe or is_30m_mfi_rising_pe

                                # Track Lower Bollinger Band touch
                                if lb_3m_pe is not None and (live_pe_ltp <= lb_3m_pe + 2.0 or (c_low_3m_pe is not None and c_low_3m_pe <= lb_3m_pe + 2.0) or p_low_15m <= lb_3m_pe + 2.0):
                                    lower_bb_touched_pe = True

                                # Middle Band Proximity & Correction at or below MB:
                                is_corrected_to_mb_pe = False
                                if mb_3m_pe is not None:
                                    is_corrected_to_mb_pe = (live_pe_ltp <= mb_3m_pe + 1.0 or (c_low_3m_pe is not None and c_low_3m_pe <= mb_3m_pe + 1.0) or p_low_15m <= mb_3m_pe + 1.0)
                                
                                # Middle Band entry condition: Correct at/below MB, bounce to open, 3m MFI rising AND any HTF MFI rising
                                is_mb_swing_entry_pe = is_corrected_to_mb_pe and is_bounce_open_pe and is_3m_mfi_rising_pe and is_htf_mfi_rising_pe and not is_swing_ob_blocked_pe

                                # If 3m MFIs are falling, force waiting for lower Bollinger Band touch
                                if is_3m_both_falling_pe:
                                    waiting_for_bb_pullback_pe = True

                                is_lower_bb_bounce_entry_pe = waiting_for_bb_pullback_pe and lower_bb_touched_pe and is_bounce_open_pe and is_3m_mfi_rising_pe and is_htf_mfi_rising_pe and not is_swing_ob_blocked_pe

                                # Retest near validated dynamic swing low
                                is_near_swing_low_pe = (p_low_15m <= dynamic_swing_low_pe + (2.5 * dynamic_tolerance_pe)) or (live_pe_ltp <= dynamic_swing_low_pe + (2.5 * dynamic_tolerance_pe))
                                is_direct_swing_entry_pe = not waiting_for_bb_pullback_pe and is_near_swing_low_pe and is_bounce_open_pe and is_3m_mfi_rising_pe and is_htf_mfi_rising_pe and not is_swing_ob_blocked_pe

                                is_swing_low_retest_entry_pe = (is_mb_swing_entry_pe or is_lower_bb_bounce_entry_pe or is_direct_swing_entry_pe) and not is_30m_both_falling_pe and not is_swing_ob_blocked_pe

                                if is_swing_low_retest_entry_pe:
                                    waiting_for_bb_pullback_pe = False
                                    lower_bb_touched_pe = False

                                # Rule #1 Entry Filter: Block fresh entry if 15m MFI(5) is at 100 / extreme overbought (>=99.0) OR MFI(14) >= 70.0 OR MFI(14) not increasing when MFI(5) >= 90
                                is_universal_ob_blocked_pe = (mfi5_15m_pe >= 99.0) or (mfi14_15m_pe >= 70.0) or (mfi5_15m_pe >= 90.0 and not (mfi14_15m_pe > prev_mfi14_15m_pe))
                                if is_universal_ob_blocked_pe:
                                    is_clean_initial_entry_pe = False
                                    is_breakout_entry_pe = False
                                    is_30m_mfi_option_pe = False
                                    is_reentry_pe = False
                                    is_recovery_reentry_pe = False
                                    is_post_breakdown_entry_pe = False
                                    is_swing_low_retest_entry_pe = False

                                # Block fresh entries after 2:45 PM (14:45 IST) for intraday safety before 15:25 square-off
                                if now_time_str >= "14:45":
                                    is_clean_initial_entry_pe = False
                                    is_breakout_entry_pe = False
                                    is_30m_mfi_option_pe = False
                                    is_reentry_pe = False
                                    is_recovery_reentry_pe = False
                                    is_post_breakdown_entry_pe = False
                                    is_swing_low_retest_entry_pe = False

                                # 9:15 Big Gap Up Retest & Bounce requirement for PE (Applies universally to ALL entry types)
                                is_pe_big_gap_up = (p_open_15m - previous_pe_high >= 30.0) or (live_pe_ltp - previous_pe_high >= 30.0)
                                if is_915_opening and is_pe_big_gap_up:
                                    is_clean_initial_entry_pe = False
                                    is_breakout_entry_pe = False
                                    is_30m_mfi_option_pe = False
                                    is_reentry_pe = False
                                    is_recovery_reentry_pe = False
                                    is_post_breakdown_entry_pe = False
                                    is_swing_low_retest_entry_pe = False
                                elif is_pe_big_gap_up and not is_915_opening:
                                    is_near_low_35pt_pe = (live_pe_ltp <= p_low_15m + 35.0)
                                    is_mfi_rising_3m_pe = (mfi5_1m_pe > prev_mfi5_1m_pe) and (mfi14_1m_pe >= prev_mfi14_1m_pe)
                                    is_extreme_oversold_reversal_pe = (mfi5_15m_pe == 0.0 and mfi14_15m_pe <= 15.0) and (live_pe_ltp >= p_low_15m + 7.0)
                                    is_bounce_to_open_pe = (live_pe_ltp >= p_open_15m) and (mfi5_1m_pe > prev_mfi5_1m_pe and mfi14_1m_pe > prev_mfi14_1m_pe)
                                    
                                    is_valid_retest_pe = is_near_low_35pt_pe and (is_mfi_rising_3m_pe or is_extreme_oversold_reversal_pe or is_bounce_to_open_pe)
                                    if not is_valid_retest_pe:
                                        is_clean_initial_entry_pe = False
                                        is_breakout_entry_pe = False
                                        is_30m_mfi_option_pe = False
                                        is_reentry_pe = False
                                        is_recovery_reentry_pe = False
                                        is_post_breakdown_entry_pe = False
                                        is_swing_low_retest_entry_pe = False

                                # --- Dynamic EPM Low Bounce LONG Entry Signal (PE) ---
                                pe_epm_low_saved = load_grid_state_epm_low("PE")
                                if pe_epm_low_saved is None or pe_epm_low_saved <= 0:
                                    pe_epm_low_saved = grid.pe_leg.epm_lower_range if (grid and grid.pe_leg) else 0.0

                                atr14_pe, stddev20_pe = calculate_atr_and_stddev(smart_api, "BFO", pe_contract.symbol_token)
                                dynamic_near_thresh_pe = max(0.25 * atr14_pe, 0.8 * stddev20_pe, 12.0)

                                is_price_near_epm_low_pe = (pe_epm_low_saved > 0.0) and (live_pe_ltp >= pe_epm_low_saved) and ((live_pe_ltp - pe_epm_low_saved) <= dynamic_near_thresh_pe)
                                prev_low_pe_check = p_low_15m if p_low_15m is not None else live_pe_ltp
                                is_bouncing_pe = (pe_epm_low_saved > 0.0) and (prev_low_pe_check <= pe_epm_low_saved + dynamic_near_thresh_pe) and (live_pe_ltp >= p_open_15m)
                                is_mfi_increasing_pe = (mfi14_15m_pe > prev_mfi14_15m_pe) or (mfi5_15m_pe == 0.0) or (prev_mfi5_15m_pe == 0.0 and mfi5_15m_pe > 0.0) or (mfi14_15m_pe <= 25.0 and mfi14_15m_pe > prev_mfi14_15m_pe)

                                is_epm_low_bounce_entry_pe = is_price_near_epm_low_pe and is_bouncing_pe and is_mfi_increasing_pe

                                if not higher_tf_block_pe:
                                    if is_epm_low_bounce_entry_pe:
                                        pe_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_pe = live_pe_ltp - 20.0
                                        entry_type_str_pe = f"PE Dynamic EPM Low Bounce LONG Entry (EPM Low: ₹{pe_epm_low_saved:.2f}, Thresh: ₹{dynamic_near_thresh_pe:.1f} | SL-20)"
                                    elif is_recovery_reentry_pe:
                                        pe_entry_signal = True
                                        recovery_reentry_eligible = False
                                        recovery_reentry_done = True
                                        lot_size = base_lot_size + 2
                                        active_sl_pe = live_pe_ltp - 20.0
                                        entry_type_str_pe = f"PE One-Time Post-SL Recovery Re-Entry (+2 Lots, Total: {lot_size} Lots | SL-20)"
                                    elif is_swing_low_retest_entry_pe:
                                        pe_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_pe = max(dynamic_swing_low_pe - 2.0, live_pe_ltp - 20.0)
                                        entry_type_str_pe = f"PE Dynamic Swing Low First Breakout Retest Entry (Pivot: ₹{dynamic_swing_low_pe:.2f}, MFI14={mfi14_15m_pe:.1f} | SL: ₹{active_sl_pe:.2f})"
                                    elif is_breakout_entry_pe:
                                        pe_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        if (mfi5_15m_pe > prev_mfi5_15m_pe and mfi14_15m_pe > prev_mfi14_15m_pe):
                                            active_sl_pe = max(p_low_15m - 15.0, live_pe_ltp - 20.0)
                                        else:
                                            active_sl_pe = live_pe_ltp - 20.0
                                        entry_type_str_pe = f"PE Previous High Breakout Momentum Entry (MFI14={mfi14_15m_pe:.1f} | SL: ₹{active_sl_pe:.2f})"
                                    elif is_direction_aligned_pe and is_clean_initial_entry_pe:
                                        pe_entry_signal = True
                                        initial_entry_happened = True
                                        lot_size = base_lot_size
                                        active_sl_pe = live_pe_ltp - 20.0
                                        entry_type_str_pe = f"PE Initial MFI Bounce (5=0 & 14<={mfi14_15m_pe:.1f} | SL-20 from Entry)"
                        
                        # Trigger CE Long Entry
                        if ce_entry_signal:
                            if bot_state == "IDLE":
                                bot_state = "CE_LONG"
                                active_contract = ce_contract
                                active_entry_price = live_ce_ltp
                                initial_entry_price = live_ce_ltp
                                staggered_scaled_in = False
                                lot_size = max(1, base_lot_size // 2) # Initial 2 Lots (40 Qty) Base Order
                                target_offset_ce = 55.0
                                active_target = active_entry_price + target_offset_ce
                                active_sl = active_sl_ce
                                entry_time = datetime.now(IST)
                                qty_to_trade = lot_size * 20
                                original_sl_distance = max(15.0, active_entry_price - active_sl)
                                trailing_active = False
                                peak_price = active_entry_price
                                offloaded = False
                                
                                sys.stdout.write("\n")
                                curr_mfi5_ce = mfi5_15m
                                prev_mfi5_ce = prev_mfi5_15m
                                curr_mfi14_ce = mfi14_15m
                                prev_mfi14_ce = prev_mfi14_15m
                                is_any_mfi_falling_at_entry_ce = (curr_mfi5_ce < prev_mfi5_ce) or (curr_mfi14_ce < prev_mfi14_ce)
                                entry_mfi_falling_15m = is_any_mfi_falling_at_entry_ce
                                logger.info("🟢 [ENTRY CE SIGNAL] CE LTP ₹%.2f triggered via %s (SL: ₹%.2f, Target: ₹%.2f) | Diagnostic: 15m MFI(5)=%.1f (prev %.1f), MFI(14)=%.1f (prev %.1f) [Falling at Entry: %s]", live_ce_ltp, entry_type_str_ce, active_sl, active_target, curr_mfi5_ce, prev_mfi5_ce, curr_mfi14_ce, prev_mfi14_ce, is_any_mfi_falling_at_entry_ce)
                                send_mobile_alert(f"🟢 *CE ENTRY SIGNAL ALIGNED ({entry_type_str_ce})*\n\n"
                                                  f"Contract: *{active_contract.trading_symbol}*\n"
                                                  f"Entry Price: ₹{active_entry_price:.2f}\n"
                                                  f"Stop Loss: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}\n"
                                                  f"Mode: *{execution_mode}* | Lot Size: *{lot_size}* ({qty_to_trade} Qty)")
                                
                                if execution_mode == "LIVE":
                                    submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_trade)
                                
                                active_strategy_name = map_entry_type_to_strategy_name(entry_type_str_ce)
                                initial_hard_sl = active_sl
                                handover_history = []
                                excel_tracker.add_order({
                                    "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                                    "mode": execution_mode,
                                    "state": "ENTRY",
                                    "trading_symbol": active_contract.trading_symbol,
                                    "price": active_entry_price,
                                    "qty": qty_to_trade,
                                    "trades_count": trades_completed + 1
                                })
                                save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                            elif bot_state == "PE_LONG":
                                if pending_swap_signal != "CE":
                                    pending_swap_signal = "CE"
                                    pending_swap_contract = ce_contract
                                    pending_swap_sl = active_sl_ce
                                    pending_swap_target = live_ce_ltp + max(15.0, grid.ce_leg.practical_target - grid.ce_leg.ltp)
                                    pending_swap_type_str = entry_type_str_ce
                                    pending_swap_time = datetime.now(IST)
                                    
                                    send_mobile_alert(
                                        f"🔔 *NEW CE SIGNAL ALIGNED ({entry_type_str_ce})*\n\n"
                                        f"Current holding: *{active_contract.trading_symbol}* (PE_LONG)\n"
                                        f"New Setup: *{ce_contract.trading_symbol}* at ₹{live_ce_ltp:.2f}\n"
                                        f"Stop Loss: ₹{pending_swap_sl:.2f} | Target: ₹{pending_swap_target:.2f}\n\n"
                                        f"👉 *Would you like to SWITCH trades?* Send 'Y' to switch instantly at Market, or 'N' to ignore."
                                    )

                        # Trigger PE Long Entry
                        elif pe_entry_signal:
                            if bot_state == "IDLE":
                                bot_state = "PE_LONG"
                                active_contract = pe_contract
                                active_entry_price = live_pe_ltp
                                initial_entry_price = live_pe_ltp
                                staggered_scaled_in = False
                                lot_size = max(1, base_lot_size // 2) # Initial 2 Lots (40 Qty) Base Order
                                target_offset_pe = 55.0
                                active_target = active_entry_price + target_offset_pe
                                active_sl = active_sl_pe
                                entry_time = datetime.now(IST)
                                qty_to_trade = lot_size * 20
                                original_sl_distance = max(15.0, active_entry_price - active_sl)
                                trailing_active = False
                                peak_price = active_entry_price
                                offloaded = False
                                
                                sys.stdout.write("\n")
                                curr_mfi5_pe = mfi5_15m_pe
                                prev_mfi5_pe = prev_mfi5_15m_pe
                                curr_mfi14_pe = mfi14_15m_pe
                                prev_mfi14_pe = prev_mfi14_15m_pe
                                is_any_mfi_falling_at_entry_pe = (curr_mfi5_pe < prev_mfi5_pe) or (curr_mfi14_pe < prev_mfi14_pe)
                                entry_mfi_falling_15m = is_any_mfi_falling_at_entry_pe
                                logger.info("🟢 [ENTRY PE SIGNAL] PE LTP ₹%.2f triggered via %s (SL: ₹%.2f, Target: ₹%.2f) | Diagnostic: 15m MFI(5)=%.1f (prev %.1f), MFI(14)=%.1f (prev %.1f) [Falling at Entry: %s]", live_pe_ltp, entry_type_str_pe, active_sl, active_target, curr_mfi5_pe, prev_mfi5_pe, curr_mfi14_pe, prev_mfi14_pe, is_any_mfi_falling_at_entry_pe)
                                send_mobile_alert(f"🟢 *PE ENTRY SIGNAL ALIGNED ({entry_type_str_pe})*\n\n"
                                                  f"Contract: *{active_contract.trading_symbol}*\n"
                                                  f"Entry Price: ₹{active_entry_price:.2f}\n"
                                                  f"Stop Loss: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}\n"
                                                  f"Mode: *{execution_mode}* | Lot Size: *{lot_size}* ({qty_to_trade} Qty)")
                                
                                if execution_mode == "LIVE":
                                    submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty_to_trade)
                                
                                active_strategy_name = map_entry_type_to_strategy_name(entry_type_str_pe)
                                initial_hard_sl = active_sl
                                handover_history = []
                                excel_tracker.add_order({
                                    "timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                                    "mode": execution_mode,
                                    "state": "ENTRY",
                                    "trading_symbol": active_contract.trading_symbol,
                                    "price": active_entry_price,
                                    "qty": qty_to_trade,
                                    "trades_count": trades_completed + 1
                                })
                                save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                            elif bot_state == "CE_LONG":
                                if pending_swap_signal != "PE":
                                    pending_swap_signal = "PE"
                                    pending_swap_contract = pe_contract
                                    pending_swap_sl = active_sl_pe
                                    pending_swap_target = live_pe_ltp + max(15.0, grid.pe_leg.practical_target - grid.pe_leg.ltp)
                                    pending_swap_type_str = entry_type_str_pe
                                    pending_swap_time = datetime.now(IST)
                                    
                                    send_mobile_alert(
                                        f"🔔 *NEW PE SIGNAL ALIGNED ({entry_type_str_pe})*\n\n"
                                        f"Current holding: *{active_contract.trading_symbol}* (CE_LONG)\n"
                                        f"New Setup: *{pe_contract.trading_symbol}* at ₹{live_pe_ltp:.2f}\n"
                                        f"Stop Loss: ₹{pending_swap_sl:.2f} | Target: ₹{pending_swap_target:.2f}\n\n"
                                        f"👉 *Would you like to SWITCH trades?* Send 'Y' to switch instantly at Market, or 'N' to ignore."
                                    )

                elif bot_state == "CE_LONG":
                    # --- CE EXIT EVALUATION & IN-FLIGHT HANDOVER ---
                    # For active position evaluation, use the actual active contract's LTP
                    live_ce_ltp = live_active_ltp
                    peak_price = max(peak_price, live_ce_ltp)
                    favorable_gain_ce = peak_price - active_entry_price
                    
                    c_open_15m, c_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", active_contract.symbol_token)
                    if c_open_15m is None:
                        c_open_15m = active_entry_price
                    if c_low_15m is None:
                        c_low_15m = min(recent_ce_low, live_ce_ltp)
                    else:
                        c_low_15m = min(c_low_15m, recent_ce_low)

                    # Pyramiding Staggered Scale-In Check CE (Add remaining 2 lots on Dip or Trend Bounce)
                    if not staggered_scaled_in:
                        is_dip_scale_in_ce = (live_ce_ltp <= initial_entry_price - 8.0) and (live_ce_ltp > active_sl) and (curr_mfi5_ce > prev_mfi5_ce)
                        is_trend_scale_in_ce = (peak_price >= initial_entry_price + 15.0) and (live_ce_ltp >= initial_entry_price + 2.0) and (curr_mfi5_ce > prev_mfi5_ce + 1.0)
                        
                        if is_dip_scale_in_ce or is_trend_scale_in_ce:
                            add_lots_ce = max(1, base_lot_size // 2)
                            scale_qty_ce = add_lots_ce * 20
                            if execution_mode == "LIVE":
                                submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", scale_qty_ce)
                            active_entry_price = (active_entry_price + live_ce_ltp) / 2.0
                            lot_size = lot_size + add_lots_ce
                            staggered_scaled_in = True
                            if is_trend_scale_in_ce:
                                active_sl = max(active_sl, initial_entry_price) # Move SL to Breakeven Cost
                            logger.info("🔥 [STAGGERED SCALE-IN CE SUCCESS] Added %d Lots at ₹%.2f | New Avg Entry: ₹%.2f | Total Lots: %d", add_lots_ce, live_ce_ltp, active_entry_price, lot_size)

                    # Fetch 15m MFIs for diagnostic & SL hold check
                    mfis_15m_ce, prev_mfis_15m_ce = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                    curr_mfi5_ce = mfis_15m_ce.get(5, 50.0)
                    curr_mfi14_ce = mfis_15m_ce.get(14, 50.0)
                    prev_mfi5_ce = prev_mfis_15m_ce.get(5, 50.0)
                    prev_mfi14_ce = prev_mfis_15m_ce.get(14, 50.0)

                    # Dynamic In-Flight Strategy Handover Evaluation (EV Maximization)
                    if ENABLE_IN_FLIGHT_HANDOVER:
                        dyn_swing_low_ce = min(c_low_15m, recent_ce_low)
                        dyn_range_ce = max(15.0, previous_ce_high - dyn_swing_low_ce)
                        dyn_tol_ce = max(3.0, min(12.0, 0.15 * dyn_range_ce))

                        ctx_ce = MarketContext(
                            timestamp=checked_at.strftime("%Y-%m-%d %H:%M:%S"),
                            open=c_open_15m,
                            high=peak_price,
                            low=c_low_15m,
                            close=live_ce_ltp,
                            volume=1.0,
                            mfi5_15m=mfi5_15m,
                            mfi14_15m=mfi14_15m,
                            prev_mfi5_15m=prev_mfi5_15m,
                            prev_mfi14_15m=prev_mfi14_15m,
                            mfi5_30m=mfi5_30m if 'mfi5_30m' in locals() else 50.0,
                            mfi14_30m=mfi14_30m if 'mfi14_30m' in locals() else 50.0,
                            prev_mfi5_30m=prev_mfi5_30m if 'prev_mfi5_30m' in locals() else 50.0,
                            prev_mfi14_30m=prev_mfi14_30m if 'prev_mfi14_30m' in locals() else 50.0,
                            prev_prev_mfi14_30m=prev_prev_mfi14_30m if 'prev_prev_mfi14_30m' in locals() else (prev_mfi14_30m if 'prev_mfi14_30m' in locals() else 50.0),
                            mfi5_60m=50.0,
                            mfi14_60m=50.0,
                            prev_mfi5_60m=50.0,
                            prev_mfi14_60m=50.0,
                            mb_20=grid.ce_leg.ltp,
                            ub_20=grid.ce_leg.target_epm,
                            lb_20=grid.ce_leg.epm_lower_range,
                            prev_high=previous_ce_high,
                            prev_low=recent_ce_low,
                            prev_close=c_open_15m,
                            recent_swing_low=recent_ce_low,
                            recent_swing_high=previous_ce_high,
                            dynamic_tolerance=dyn_tol_ce,
                            is_0915_bar=False,
                            is_big_gap_up=False,
                            allow_reentry=allow_reentry_live,
                            recovery_eligible=recovery_reentry_eligible,
                            initial_entry_done=initial_entry_happened
                        )

                        pos_ce = Position(
                            trade_id=trades_completed + 1,
                            strategy_name=active_strategy_name or "Initial Dual MFI Lower Band Bounce",
                            side=PositionSide.LONG,
                            entry_time=entry_time.strftime("%Y-%m-%d %H:%M:%S") if entry_time else "",
                            entry_price=active_entry_price,
                            current_sl=active_sl,
                            initial_sl=initial_hard_sl if initial_hard_sl > 0 else (active_entry_price - 20.0),
                            target_price=active_target,
                            peak_price=peak_price,
                            lot_size=lot_size,
                            handover_history=handover_history
                        )

                        handover_decision = handover_engine._evaluate_in_flight_handover(ctx_ce, pos_ce)
                        if handover_decision.switch_approved:
                            from_strat = active_strategy_name or "Initial Dual MFI Lower Band Bounce"
                            to_strat = handover_decision.to_strategy
                            active_strategy_name = to_strat
                            active_target = handover_decision.adjusted_target
                            
                            # GLOBAL RISK INVARIANT: Handover is NEVER permitted to widen the initial hard stop loss
                            guaranteed_sl = max(initial_hard_sl if initial_hard_sl > 0 else active_entry_price - 20.0, active_sl, handover_decision.adjusted_sl)
                            active_sl = max(active_sl, guaranteed_sl)
                            
                            ev_delta = handover_decision.projected_ev - handover_decision.current_ev
                            handover_record = {
                                "time": ctx_ce.timestamp,
                                "from_strategy": from_strat,
                                "to_strategy": to_strat,
                                "current_ev": round(handover_decision.current_ev, 2),
                                "projected_ev": round(handover_decision.projected_ev, 2),
                                "delta_ev": round(ev_delta, 2),
                                "adjusted_sl": round(active_sl, 2),
                                "adjusted_target": round(active_target, 2),
                                "reason": handover_decision.reason
                            }
                            handover_history.append(handover_record)
                            entry_type_str_ce = f"CE {to_strat} (Handover from {from_strat})"
                            
                            sys.stdout.write("\n")
                            logger.info("🔄 [STRATEGY HANDOVER] Switched from '%s' to '%s' | EV: %.1f -> %.1f (ΔEV: +%.1f pts) | New Target: ₹%.2f | Trailing SL: ₹%.2f | Reason: %s",
                                        from_strat, to_strat, handover_decision.current_ev, handover_decision.projected_ev, ev_delta, active_target, active_sl, handover_decision.reason)
                            
                            send_mobile_alert(
                                f"🔄 *[STRATEGY HANDOVER]*\n\n"
                                f"Contract: *{active_contract.trading_symbol}* (CE_LONG)\n"
                                f"Switched: *{from_strat}* ➡️ *{to_strat}*\n\n"
                                f"📈 *EV Delta:* +{ev_delta:.1f} pts ({handover_decision.current_ev:.1f} ➡️ {handover_decision.projected_ev:.1f})\n"
                                f"🎯 *New Target:* ₹{active_target:.2f}\n"
                                f"🛡️ *Trailing SL:* ₹{active_sl:.2f}\n"
                                f"📌 *Reason:* {handover_decision.reason}"
                            )
                            save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                    
                    # Check Overbought / Upper BB Proximity
                    is_overbought_or_near_ub_ce = (curr_mfi5_ce >= 80.0 or prev_mfi5_ce >= 80.0 or curr_mfi14_ce >= 70.0) or (peak_price >= grid.ce_leg.target_epm - 10.0 or live_ce_ltp >= grid.ce_leg.target_epm - 10.0)
                    mfi_rising_ce = ((curr_mfi5_ce > prev_mfi5_ce) or (curr_mfi14_ce > prev_mfi14_ce)) and not is_overbought_or_near_ub_ce

                    # Multi-Stage Profit Locking & Trailing SL:
                    # User Rule: Multi-stage profit locking should start above 40 points (+10 pts -> Entry+3, +18 pts -> Entry+10).
                    # No trail before 40 points as long as MFI(14) rising in 15 minutes.
                    is_mfi14_rising_15m_ce = (curr_mfi14_ce > prev_mfi14_ce) or (curr_mfi14_ce >= prev_mfi14_ce and curr_mfi5_ce > prev_mfi5_ce)

                    # Institutional Dual-Regime Trend Engine:
                    # 1. Strong Trend Regime (MFI14 >= 50 or HTF Rising): Allow MB dips for big swing highs
                    # 2. Exhaustion / Sideways Regime (Overbought / Retest Failure): Tight Smart Trailing
                    is_strong_institutional_trend_ce = (is_mfi14_rising_15m_ce or curr_mfi14_ce >= 50.0) and (mfi14_30m >= prev_mfi14_30m)

                    if is_strong_institutional_trend_ce:
                        # Allow price to breathe/dip to Middle Band; trail only after major +50pt surge
                        if favorable_gain_ce >= 70.0:
                            active_sl = max(active_sl, peak_price - 20.0) # Lock major runner gains
                        elif favorable_gain_ce >= 50.0:
                            active_sl = max(active_sl, active_entry_price + 20.0)
                        elif favorable_gain_ce >= 30.0:
                            active_sl = max(active_sl, active_entry_price) # Move to Cost Price / Breakeven
                        else:
                            active_sl = max(active_sl, active_entry_price - 20.0) # Hold risk floor on MB dips
                    else:
                        # Sideways / Exhaustion Regime: Active Smart Trailing
                        if favorable_gain_ce >= 40.0:
                            active_sl = max(active_sl, active_entry_price + 15.0)
                        elif favorable_gain_ce >= 20.0:
                            active_sl = max(active_sl, active_entry_price + 3.0)
                        elif favorable_gain_ce >= 12.0:
                            active_sl = max(active_sl, active_entry_price - 5.0)

                    # Calculate 3X risk-reward Take Profit target based on original risk distance
                    surge_target_price = active_entry_price + (3 * original_sl_distance)
                    
                    elapsed_mins = (datetime.now(IST) - entry_time).total_seconds() / 60.0
                    is_surge_window = (elapsed_mins <= 60.0)  # Within 1-2 30-min candles (60 mins)
                    
                    # Check for Dual MFI Exit conditions
                    mfis_15m_ce, prev_mfis_15m_ce = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                    curr_mfi5_ce = mfis_15m_ce.get(5, 50.0)
                    curr_mfi14_ce = mfis_15m_ce.get(14, 50.0)
                    prev_mfi5_ce = prev_mfis_15m_ce.get(5, 50.0)
                    prev_mfi14_ce = prev_mfis_15m_ce.get(14, 50.0)

                    # Exit Rules & Trend Riding Logic:
                    is_30m_htf_trend_rising_ce = (mfi5_30m >= prev_mfi5_30m and mfi14_30m >= prev_mfi14_30m)
                    is_price_and_mfi5_rising_ce = (live_ce_ltp > active_entry_price and curr_mfi5_ce > prev_mfi5_ce)
                    is_htf_hold_trend_ce = is_30m_htf_trend_rising_ce or is_price_and_mfi5_rising_ce

                    is_30m_dual_mfi_fall_exit_ce = (mfi5_30m < prev_mfi5_30m and mfi14_30m < prev_mfi14_30m) and (live_ce_ltp > active_entry_price)
                    is_ub_reached_dual_mfi_fall_ce = (peak_price >= grid.ce_leg.ltp + 15.0 or live_ce_ltp >= grid.ce_leg.ltp + 15.0) and (curr_mfi5_ce < prev_mfi5_ce and curr_mfi14_ce < prev_mfi14_ce)

                    # 1. Middle Band Rejection if MFI(5) or MFI(14) is falling (within 5 pts from peak)
                    is_mb_rejection_ce = (peak_price >= grid.ce_leg.ltp - 5.0 and (curr_mfi5_ce < prev_mfi5_ce or curr_mfi14_ce < prev_mfi14_ce) and live_ce_ltp <= peak_price - 5.0) and not is_htf_hold_trend_ce
                    # 2. Exit if BOTH MFI(5) and MFI(14) are falling at opening/close while in profit (override if HTF rising)
                    is_dual_mfi_falling_ce = (curr_mfi5_ce < prev_mfi5_ce) and (curr_mfi14_ce < prev_mfi14_ce) and (live_ce_ltp > active_entry_price) and not is_htf_hold_trend_ce
                    # 3. Hold trend while MFI(14) is rising; exit when MFI(14) falls after overbought
                    is_overbought_mfi14_fall_ce = (curr_mfi5_ce >= 95.0 or prev_mfi5_ce >= 95.0) and (curr_mfi14_ce < prev_mfi14_ce)

                    # Rule #1 Specific Exit: For "MB Consolidation Re-Entry" / "MFI(14) Trend Re-Entry"
                    is_reentry_active_ce = ("MB Consolidation" in entry_type_str_ce if 'entry_type_str_ce' in locals() else True)
                    is_near_or_above_ub_ce = (peak_price >= grid.ce_leg.target_epm - 5.0 or live_ce_ltp >= grid.ce_leg.target_epm - 5.0)
                    is_any_mfi_falling_ce = (curr_mfi5_ce < prev_mfi5_ce or curr_mfi14_ce < prev_mfi14_ce)
                    is_both_mfi_falling_ce = (curr_mfi5_ce < prev_mfi5_ce and curr_mfi14_ce < prev_mfi14_ce)

                    # 1. Strict Exit when both 15 min MFIs fall rather than waiting for SL
                    is_reentry_dual_mfi_fall_exit_ce = is_reentry_active_ce and is_both_mfi_falling_ce

                    # 2. Strict Exit if MFI(14) or both MFIs falling in 15 min when price already near Upper Bollinger Band
                    is_reentry_ub_mfi14_or_dual_fall_ce = is_reentry_active_ce and is_near_or_above_ub_ce and (curr_mfi14_ce < prev_mfi14_ce or is_both_mfi_falling_ce)

                    is_reentry_ub_cross_fall_ce = is_reentry_dual_mfi_fall_exit_ce or is_reentry_ub_mfi14_or_dual_fall_ce

                    # Rule #3 Specific Exit: For Breakout Entry - Multi-factor breakout exit rules
                    is_breakout_trade_ce = ("Breakout" in entry_type_str_ce if 'entry_type_str_ce' in locals() else False)
                    is_breakout_mfi14_or_both_fall_ce = is_breakout_trade_ce and (curr_mfi14_ce < prev_mfi14_ce or is_both_mfi_falling_ce)
                    is_breakout_weak_close_retrace_ce = is_breakout_trade_ce and (peak_price >= active_entry_price + 5.0) and (live_ce_ltp <= c_open_15m or peak_price - live_ce_ltp >= 6.0) and not (curr_mfi14_ce > prev_mfi14_ce)
                    is_breakout_mb_rejection_ce = is_breakout_trade_ce and (peak_price >= grid.ce_leg.ltp - 5.0 and peak_price <= grid.ce_leg.ltp + 8.0) and (curr_mfi14_ce < prev_mfi14_ce or curr_mfi5_ce >= 80.0 or curr_mfi14_ce >= 68.0) and (live_ce_ltp <= peak_price - 3.0)
                    is_breakout_mfi100_fall_ce = is_breakout_trade_ce and (curr_mfi5_ce >= 100.0 and curr_mfi14_ce < prev_mfi14_ce)
                    is_breakout_3m_fall_ce = is_breakout_trade_ce and (curr_mfi5_ce >= 98.0) and (prev_mfi5_ce - curr_mfi5_ce > 1.0)

                    # Rule #5 Specific Exit: For Recovery Re-Entry - Mandatory exit on 15m dual MFI fall
                    is_recovery_mfi_fall_ce = ("Recovery Re-Entry" in entry_type_str_ce if 'entry_type_str_ce' in locals() else False) and (curr_mfi5_ce < prev_mfi5_ce and curr_mfi14_ce < prev_mfi14_ce)

                    # Rule #6 Specific Exit: For Post-Breakdown Oversold Bounce - Exit on dual MFI fall or MFI14 fall with price rejection
                    is_post_breakdown_rejection_mfi_fall_ce = ("Post-Breakdown" in entry_type_str_ce if 'entry_type_str_ce' in locals() else False) and ((curr_mfi5_ce < prev_mfi5_ce and curr_mfi14_ce < prev_mfi14_ce) or ((curr_mfi14_ce < prev_mfi14_ce) and (live_ce_ltp < c_open_15m or live_ce_ltp <= peak_price - 3.0)))

                    if ("Post-Breakdown" in entry_type_str_ce if 'entry_type_str_ce' in locals() else False) and (curr_mfi5_ce > prev_mfi5_ce and curr_mfi14_ce > prev_mfi14_ce):
                        active_sl = max(active_sl, active_entry_price - 20.0)

                    # Rule: Exit on same candle closing if 15 min any MFI was falling at entry time and price crossed Upper BB with MFI falling
                    res_bb_ce_check = get_3m_bollinger_bands(smart_api, "BFO", active_contract.symbol_token)
                    ub_price_ce = res_bb_ce_check[1] if (res_bb_ce_check and res_bb_ce_check[1]) else grid.ce_leg.target_epm
                    is_ub_crossed_ce = (peak_price >= ub_price_ce) or (live_ce_ltp >= ub_price_ce)
                    is_mfi_falling_now_ce = (curr_mfi5_ce < prev_mfi5_ce) or (curr_mfi14_ce < prev_mfi14_ce)
                    is_same_candle_ub_mfi_fall_exit_ce = entry_mfi_falling_15m and is_ub_crossed_ce and is_mfi_falling_now_ce

                    # Swing Low Breakout / Retest Strategy Specific Exit Rule:
                    # User Rule: Wait after entry if MFI(14) is rising till Upper Bollinger band price rejection or MFI down near or above Upper Band
                    is_swing_trade_ce = ("Swing Low" in entry_type_str_ce if 'entry_type_str_ce' in locals() else False) or (active_strategy_name == "Dynamic Swing Low First Breakout Retest Entry")
                    is_swing_mfi14_rising_ce = (curr_mfi14_ce > prev_mfi14_ce) or (curr_mfi14_ce >= prev_mfi14_ce and curr_mfi5_ce > prev_mfi5_ce)
                    is_swing_ub_rejection_ce = (live_ce_ltp >= ub_price_ce - 2.0 or peak_price >= ub_price_ce - 2.0) and (live_ce_ltp <= c_open_15m or peak_price - live_ce_ltp >= 5.0)
                    is_swing_mfi_down_near_ub_ce = (peak_price >= ub_price_ce - 5.0 or live_ce_ltp >= ub_price_ce - 5.0 or curr_mfi14_ce >= 65.0) and (curr_mfi14_ce < prev_mfi14_ce or curr_mfi5_ce < prev_mfi5_ce)
                    is_swing_exit_ce = is_swing_trade_ce and ((is_swing_mfi14_rising_ce and (is_swing_ub_rejection_ce or is_swing_mfi_down_near_ub_ce) and (live_ce_ltp >= active_entry_price + 8.0 or peak_price >= active_entry_price + 15.0)) or (not is_swing_mfi14_rising_ce and is_both_mfi_falling_ce))

                    # Extreme Overbought High Rejection Exit Rule:
                    is_overbought_high_rejection_ce = (curr_mfi14_ce >= 70.0 or curr_mfi5_ce >= 80.0) and (live_ce_ltp >= previous_ce_high or peak_price >= previous_ce_high) and (live_ce_ltp < previous_ce_high or live_ce_ltp <= c_open_15m or peak_price - live_ce_ltp >= 5.0)

                    # Profit Booking Exit Logic (80+ points OR MFI(14) or both MFI falling in 15m/3m frame)
                    points_gained_ce = live_ce_ltp - active_entry_price
                    is_80pt_profit_booking_ce = (points_gained_ce >= 80.0)
                    mfi_falling_15m_ce = (curr_mfi14_ce < prev_mfi14_ce) or (curr_mfi5_ce < prev_mfi5_ce and curr_mfi14_ce < prev_mfi14_ce)
                    mfi_falling_3m_ce = (mfi14_3m_ce < prev_mfi14_3m_ce) or (mfi5_3m_ce < prev_mfi5_3m_ce and mfi14_3m_ce < prev_mfi14_3m_ce) if ('mfi14_3m_ce' in locals() and 'prev_mfi14_3m_ce' in locals()) else False
                    is_mfi_falling_profit_booking_ce = (points_gained_ce > 0.0) and (mfi_falling_15m_ce or mfi_falling_3m_ce)
                    is_profit_booking_exit_ce = is_80pt_profit_booking_ce or is_mfi_falling_profit_booking_ce

                    # Same Day EOD Mandatory Exit (3:25 PM / 3:30 PM cutoff)
                    now_time_str_exit = datetime.now(IST).strftime("%H:%M")
                    is_eod_exit_live = (now_time_str_exit >= "15:25")

                    is_mfi_exit_triggered_ce = is_profit_booking_exit_ce or is_eod_exit_live or is_swing_exit_ce or is_overbought_high_rejection_ce or is_breakout_mfi14_or_both_fall_ce or is_breakout_weak_close_retrace_ce or is_breakout_mb_rejection_ce or is_30m_dual_mfi_fall_exit_ce or is_ub_reached_dual_mfi_fall_ce or is_mb_rejection_ce or is_dual_mfi_falling_ce or is_overbought_mfi14_fall_ce or is_reentry_ub_cross_fall_ce or is_breakout_mfi100_fall_ce or is_breakout_3m_fall_ce or is_recovery_mfi_fall_ce or is_post_breakdown_rejection_mfi_fall_ce or is_same_candle_ub_mfi_fall_exit_ce
                    
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
                        if exit_state_str == "EXIT_SL":
                            ce_sl_hit_today = True
                        
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
                        active_strategy_name = ""
                        initial_hard_sl = 0.0
                        handover_history = []
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        # If Hard SL hit (not TSL), allow One-Time Post-SL Recovery Re-Entry
                        if not trailing_active and not recovery_reentry_done:
                            recovery_reentry_eligible = True
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                    # 3. Standard Practical Target exit or MFI / 3m Upper BB Target Exit
                    if bot_state != "CE_LONG" or active_contract is None:
                        continue
                    c_mfi_val, c_mfi_prev_val, _ = get_15m_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    c_1h_mfi, p_1h_mfi = get_1h_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    res_bb_ce = get_3m_bollinger_bands(smart_api, "BFO", active_contract.symbol_token)
                    ub_3m_ce = res_bb_ce[1] if res_bb_ce else None
                    
                    # Fetch 3m MFIs (5 and 14) for 3m Upper BB momentum evaluation
                    mfis_3m_ce, prev_mfis_3m_ce = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "THREE_MINUTE", [5, 14])
                    curr_mfi5_3m_ce = mfis_3m_ce.get(5, 50.0)
                    prev_mfi5_3m_ce = prev_mfis_3m_ce.get(5, 50.0)
                    curr_mfi14_3m_ce = mfis_3m_ce.get(14, 50.0)
                    prev_mfi14_3m_ce = prev_mfis_3m_ce.get(14, 50.0)

                    is_1h_mfi_falling = (c_1h_mfi < p_1h_mfi)
                    is_3m_ub_near = (ub_3m_ce is not None) and (live_ce_ltp >= ub_3m_ce - 5.0)
                    is_40pt_gain = (favorable_gain_ce >= 40.0)

                    # Both 3m MFIs increasing (by at least 1.0 point)
                    both_mfi_increasing_3m = (curr_mfi5_3m_ce >= prev_mfi5_3m_ce + 1.0) and (curr_mfi14_3m_ce >= prev_mfi14_3m_ce + 1.0)
                    # Hold while MFI(14) is rising after MFI(5) reaches 100
                    mfi14_rising_after_100 = (curr_mfi5_3m_ce >= 99.0 or c_mfi_val >= 99.0) and (curr_mfi14_3m_ce >= prev_mfi14_3m_ce)
                    hold_due_to_surging_mfi = both_mfi_increasing_3m or mfi14_rising_after_100

                    # Exit allowed if MFI(14) is falling OR both 3m MFIs are falling
                    is_mfi14_falling_3m = (curr_mfi14_3m_ce < prev_mfi14_3m_ce)
                    both_mfi_falling_3m = (curr_mfi5_3m_ce < prev_mfi5_3m_ce) and is_mfi14_falling_3m
                    exit_mfi_confirmed = is_mfi14_falling_3m or both_mfi_falling_3m

                    # Higher Time Frame MFI Hierarchy: follow 3m Upper BB / +40pt profit booking ONLY IF any HTF MFI is falling
                    is_htf_mfi_falling = is_1h_mfi_falling or (c_mfi_val < c_mfi_prev_val)

                    # Book profit at 3m Upper BB or +40pt gain ONLY IF exit is confirmed by falling 3m MFI AND any HTF MFI is falling
                    is_3m_bb_profit_booking = (is_3m_ub_near or is_40pt_gain) and exit_mfi_confirmed and is_htf_mfi_falling and not hold_due_to_surging_mfi
                    
                    is_mfi_tp_hit = (c_mfi_val >= 99.0 and not (curr_mfi14_3m_ce >= prev_mfi14_3m_ce)) or (live_ce_ltp >= active_entry_price + 100.0 and c_mfi_val < c_mfi_prev_val) or is_3m_bb_profit_booking
                    
                    if not trailing_active and (live_ce_ltp >= active_target or is_mfi_tp_hit):
                        tp_reason = "3m Upper BB / +40pt Profit Booking (HTF MFI Falling)" if is_3m_bb_profit_booking else ("MFI Target (100 / 100pt+ & Declining)" if is_mfi_tp_hit else f"Practical Target (₹{active_target:.2f})")
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
                        active_strategy_name = ""
                        initial_hard_sl = 0.0
                        handover_history = []
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                elif bot_state == "PE_LONG":
                    # --- PE EXIT EVALUATION & IN-FLIGHT HANDOVER ---
                    # For active position evaluation, use the actual active contract's LTP
                    live_pe_ltp = live_active_ltp
                    peak_price = max(peak_price, live_pe_ltp)
                    favorable_gain_pe = peak_price - active_entry_price
                    
                    p_open_15m, p_low_15m = get_current_15m_candle_ohl(smart_api, "BFO", active_contract.symbol_token)
                    if p_open_15m is None:
                        p_open_15m = active_entry_price
                    if p_low_15m is None:
                        p_low_15m = min(recent_pe_low, live_pe_ltp)
                    else:
                        p_low_15m = min(p_low_15m, recent_pe_low)

                    # Pyramiding Staggered Scale-In Check PE (Add remaining 2 lots on Dip or Trend Bounce)
                    if not staggered_scaled_in:
                        is_dip_scale_in_pe = (live_pe_ltp <= initial_entry_price - 8.0) and (live_pe_ltp > active_sl) and (curr_mfi5_pe > prev_mfi5_pe)
                        is_trend_scale_in_pe = (peak_price >= initial_entry_price + 15.0) and (live_pe_ltp >= initial_entry_price + 2.0) and (curr_mfi5_pe > prev_mfi5_pe + 1.0)
                        
                        if is_dip_scale_in_pe or is_trend_scale_in_pe:
                            add_lots_pe = max(1, base_lot_size // 2)
                            scale_qty_pe = add_lots_pe * 20
                            if execution_mode == "LIVE":
                                submit_angel_order(smart_api, active_contract.trading_symbol, active_contract.symbol_token, "BUY", scale_qty_pe)
                            active_entry_price = (active_entry_price + live_pe_ltp) / 2.0
                            lot_size = lot_size + add_lots_pe
                            staggered_scaled_in = True
                            if is_trend_scale_in_pe:
                                active_sl = max(active_sl, initial_entry_price) # Move SL to Breakeven Cost
                            logger.info("🔥 [STAGGERED SCALE-IN PE SUCCESS] Added %d Lots at ₹%.2f | New Avg Entry: ₹%.2f | Total Lots: %d", add_lots_pe, live_pe_ltp, active_entry_price, lot_size)

                    # Fetch 15m MFIs for diagnostic & SL hold check
                    mfis_15m_pe_act, prev_mfis_15m_pe_act = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                    curr_mfi5_pe = mfis_15m_pe_act.get(5, 50.0)
                    curr_mfi14_pe = mfis_15m_pe_act.get(14, 50.0)
                    prev_mfi5_pe = prev_mfis_15m_pe_act.get(5, 50.0)
                    prev_mfi14_pe = prev_mfis_15m_pe_act.get(14, 50.0)

                    # Dynamic In-Flight Strategy Handover Evaluation (EV Maximization)
                    if ENABLE_IN_FLIGHT_HANDOVER:
                        dyn_swing_low_pe = min(p_low_15m, recent_pe_low)
                        dyn_range_pe = max(15.0, previous_pe_high - dyn_swing_low_pe)
                        dyn_tol_pe = max(3.0, min(12.0, 0.15 * dyn_range_pe))

                        ctx_pe = MarketContext(
                            timestamp=checked_at.strftime("%Y-%m-%d %H:%M:%S"),
                            open=p_open_15m,
                            high=peak_price,
                            low=p_low_15m,
                            close=live_pe_ltp,
                            volume=1.0,
                            mfi5_15m=mfi5_15m_pe,
                            mfi14_15m=mfi14_15m_pe,
                            prev_mfi5_15m=prev_mfi5_15m_pe,
                            prev_mfi14_15m=prev_mfi14_15m_pe,
                            mfi5_30m=mfi5_30m_pe if 'mfi5_30m_pe' in locals() else (mfi5_30m if 'mfi5_30m' in locals() else 50.0),
                            mfi14_30m=mfi14_30m_pe if 'mfi14_30m_pe' in locals() else (mfi14_30m if 'mfi14_30m' in locals() else 50.0),
                            prev_mfi5_30m=prev_mfi5_30m_pe if 'prev_mfi5_30m_pe' in locals() else (prev_mfi5_30m if 'prev_mfi5_30m' in locals() else 50.0),
                            prev_mfi14_30m=prev_mfi14_30m_pe if 'prev_mfi14_30m_pe' in locals() else (prev_mfi14_30m if 'prev_mfi14_30m' in locals() else 50.0),
                            prev_prev_mfi14_30m=prev_prev_mfi14_30m_pe if 'prev_prev_mfi14_30m_pe' in locals() else (prev_mfi14_30m_pe if 'prev_mfi14_30m_pe' in locals() else 50.0),
                            mfi5_60m=50.0,
                            mfi14_60m=50.0,
                            prev_mfi5_60m=50.0,
                            prev_mfi14_60m=50.0,
                            mb_20=grid.pe_leg.ltp,
                            ub_20=grid.pe_leg.target_epm,
                            lb_20=grid.pe_leg.epm_lower_range,
                            prev_high=previous_pe_high,
                            prev_low=recent_pe_low,
                            prev_close=p_open_15m,
                            recent_swing_low=recent_pe_low,
                            recent_swing_high=previous_pe_high,
                            dynamic_tolerance=dyn_tol_pe,
                            is_0915_bar=False,
                            is_big_gap_up=False,
                            allow_reentry=allow_reentry_live,
                            recovery_eligible=recovery_reentry_eligible,
                            initial_entry_done=initial_entry_happened
                        )

                        pos_pe = Position(
                            trade_id=trades_completed + 1,
                            strategy_name=active_strategy_name or "Initial Dual MFI Lower Band Bounce",
                            side=PositionSide.LONG,
                            entry_time=entry_time.strftime("%Y-%m-%d %H:%M:%S") if entry_time else "",
                            entry_price=active_entry_price,
                            current_sl=active_sl,
                            initial_sl=initial_hard_sl if initial_hard_sl > 0 else (active_entry_price - 20.0),
                            target_price=active_target,
                            peak_price=peak_price,
                            lot_size=lot_size,
                            handover_history=handover_history
                        )

                        handover_decision_pe = handover_engine._evaluate_in_flight_handover(ctx_pe, pos_pe)
                        if handover_decision_pe.switch_approved:
                            from_strat = active_strategy_name or "Initial Dual MFI Lower Band Bounce"
                            to_strat = handover_decision_pe.to_strategy
                            active_strategy_name = to_strat
                            active_target = handover_decision_pe.adjusted_target
                            
                            # GLOBAL RISK INVARIANT: Handover is NEVER permitted to widen the initial hard stop loss
                            guaranteed_sl = max(initial_hard_sl if initial_hard_sl > 0 else active_entry_price - 20.0, active_sl, handover_decision_pe.adjusted_sl)
                            active_sl = max(active_sl, guaranteed_sl)
                            
                            ev_delta = handover_decision_pe.projected_ev - handover_decision_pe.current_ev
                            handover_record = {
                                "time": ctx_pe.timestamp,
                                "from_strategy": from_strat,
                                "to_strategy": to_strat,
                                "current_ev": round(handover_decision_pe.current_ev, 2),
                                "projected_ev": round(handover_decision_pe.projected_ev, 2),
                                "delta_ev": round(ev_delta, 2),
                                "adjusted_sl": round(active_sl, 2),
                                "adjusted_target": round(active_target, 2),
                                "reason": handover_decision_pe.reason
                            }
                            handover_history.append(handover_record)
                            entry_type_str_pe = f"PE {to_strat} (Handover from {from_strat})"
                            
                            sys.stdout.write("\n")
                            logger.info("🔄 [STRATEGY HANDOVER] Switched from '%s' to '%s' | EV: %.1f -> %.1f (ΔEV: +%.1f pts) | New Target: ₹%.2f | Trailing SL: ₹%.2f | Reason: %s",
                                        from_strat, to_strat, handover_decision_pe.current_ev, handover_decision_pe.projected_ev, ev_delta, active_target, active_sl, handover_decision_pe.reason)
                            
                            send_mobile_alert(
                                f"🔄 *[STRATEGY HANDOVER]*\n\n"
                                f"Contract: *{active_contract.trading_symbol}* (PE_LONG)\n"
                                f"Switched: *{from_strat}* ➡️ *{to_strat}*\n\n"
                                f"📈 *EV Delta:* +{ev_delta:.1f} pts ({handover_decision_pe.current_ev:.1f} ➡️ {handover_decision_pe.projected_ev:.1f})\n"
                                f"🎯 *New Target:* ₹{active_target:.2f}\n"
                                f"🛡️ *Trailing SL:* ₹{active_sl:.2f}\n"
                                f"📌 *Reason:* {handover_decision_pe.reason}"
                            )
                            save_bot_memory_full(trades_completed, current_slot, grid, ce_contract, pe_contract, bot_state, active_contract, active_entry_price, active_sl, active_target, peak_price, trailing_active, offloaded, lot_size, entry_time, entry_mfi_falling_15m, active_strategy_name, initial_hard_sl, handover_history)
                    
                    # Check Overbought / Upper BB Proximity
                    is_overbought_or_near_ub_pe = (curr_mfi5_pe >= 80.0 or prev_mfi5_pe >= 80.0 or curr_mfi14_pe >= 70.0) or (peak_price >= grid.pe_leg.target_epm - 10.0 or live_pe_ltp >= grid.pe_leg.target_epm - 10.0)
                    mfi_rising_pe = ((curr_mfi5_pe > prev_mfi5_pe) or (curr_mfi14_pe > prev_mfi14_pe)) and not is_overbought_or_near_ub_pe

                    # Multi-Stage Profit Locking & Trailing SL PE:
                    # User Rule: Multi-stage profit locking should start above 40 points (+10 pts -> Entry+3, +18 pts -> Entry+10).
                    # No trail before 40 points as long as MFI(14) rising in 15 minutes.
                    is_mfi14_rising_15m_pe = (curr_mfi14_pe > prev_mfi14_pe) or (curr_mfi14_pe >= prev_mfi14_pe and curr_mfi5_pe > prev_mfi5_pe)

                    # Institutional Dual-Regime Trend Engine PE:
                    # 1. Strong Trend Regime (MFI14 >= 50 or HTF Rising): Allow MB dips for big swing highs
                    # 2. Exhaustion / Sideways Regime (Overbought / Retest Failure): Tight Smart Trailing
                    is_strong_institutional_trend_pe = (is_mfi14_rising_15m_pe or curr_mfi14_pe >= 50.0) and (mfi14_30m_pe >= prev_mfi14_30m_pe)

                    if is_strong_institutional_trend_pe:
                        # Allow price to breathe/dip to Middle Band; trail only after major +50pt surge
                        if favorable_gain_pe >= 70.0:
                            active_sl = max(active_sl, peak_price - 20.0) # Lock major runner gains
                        elif favorable_gain_pe >= 50.0:
                            active_sl = max(active_sl, active_entry_price + 20.0)
                        elif favorable_gain_pe >= 30.0:
                            active_sl = max(active_sl, active_entry_price) # Move to Cost Price / Breakeven
                        else:
                            active_sl = max(active_sl, active_entry_price - 20.0) # Hold risk floor on MB dips
                    else:
                        # Sideways / Exhaustion Regime: Active Smart Trailing
                        if favorable_gain_pe >= 40.0:
                            active_sl = max(active_sl, active_entry_price + 15.0)
                        elif favorable_gain_pe >= 20.0:
                            active_sl = max(active_sl, active_entry_price + 3.0)
                        elif favorable_gain_pe >= 12.0:
                            active_sl = max(active_sl, active_entry_price - 5.0)

                    # Calculate 3X risk-reward Take Profit target based on original risk distance
                    surge_target_price = active_entry_price + (3 * original_sl_distance)
                    
                    elapsed_mins = (datetime.now(IST) - entry_time).total_seconds() / 60.0
                    is_surge_window = (elapsed_mins <= 60.0)  # Within 1-2 30-min candles (60 mins)
                    
                    # Check for Dual MFI Exit conditions
                    mfis_15m_pe_act, prev_mfis_15m_pe_act = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14])
                    curr_mfi5_pe = mfis_15m_pe_act.get(5, 50.0)
                    curr_mfi14_pe = mfis_15m_pe_act.get(14, 50.0)
                    prev_mfi5_pe = prev_mfis_15m_pe_act.get(5, 50.0)
                    prev_mfi14_pe = prev_mfis_15m_pe_act.get(14, 50.0)

                    # Exit Rules:
                    # 1. Middle Band Rejection if MFI(5) or MFI(14) is falling (within 5 pts from peak)
                    is_mb_rejection_pe = (peak_price >= grid.pe_leg.ltp - 5.0 and (curr_mfi5_pe < prev_mfi5_pe or curr_mfi14_pe < prev_mfi14_pe) and live_pe_ltp <= peak_price - 5.0)
                    # 2. Exit if BOTH MFI(5) and MFI(14) are falling at opening/close while in profit
                    is_dual_mfi_falling_pe = (curr_mfi5_pe < prev_mfi5_pe) and (curr_mfi14_pe < prev_mfi14_pe) and (live_pe_ltp > active_entry_price)
                    # 3. Hold trend while MFI(14) is rising; exit when MFI(14) falls after overbought
                    is_overbought_mfi14_fall_pe = (curr_mfi5_pe >= 95.0 or prev_mfi5_pe >= 95.0) and (curr_mfi14_pe < prev_mfi14_pe)

                    # Rule #1 Specific Exit: For "MB Consolidation Re-Entry" / "MFI(14) Trend Re-Entry" PE
                    is_reentry_active_pe = ("MB Consolidation" in entry_type_str_pe if 'entry_type_str_pe' in locals() else True)
                    is_near_or_above_ub_pe = (peak_price >= grid.pe_leg.target_epm - 5.0 or live_pe_ltp >= grid.pe_leg.target_epm - 5.0)
                    is_any_mfi_falling_pe = (curr_mfi5_pe < prev_mfi5_pe or curr_mfi14_pe < prev_mfi14_pe)
                    is_both_mfi_falling_pe = (curr_mfi5_pe < prev_mfi5_pe and curr_mfi14_pe < prev_mfi14_pe)

                    # 1. Strict Exit when both 15 min MFIs fall rather than waiting for SL
                    is_reentry_dual_mfi_fall_exit_pe = is_reentry_active_pe and is_both_mfi_falling_pe

                    # 2. Strict Exit if MFI(14) or both MFIs falling in 15 min when price already near Upper Bollinger Band
                    is_reentry_ub_mfi14_or_dual_fall_pe = is_reentry_active_pe and is_near_or_above_ub_pe and (curr_mfi14_pe < prev_mfi14_pe or is_both_mfi_falling_pe)

                    is_reentry_ub_cross_fall_pe = is_reentry_dual_mfi_fall_exit_pe or is_reentry_ub_mfi14_or_dual_fall_pe

                    # Rule #3 Specific Exit: For Breakout Entry PE - Multi-factor breakout exit rules
                    is_breakout_trade_pe = ("Breakout" in entry_type_str_pe if 'entry_type_str_pe' in locals() else False)
                    is_breakout_mfi14_or_both_fall_pe = is_breakout_trade_pe and (curr_mfi14_pe < prev_mfi14_pe or is_both_mfi_falling_pe)
                    is_breakout_weak_close_retrace_pe = is_breakout_trade_pe and (peak_price >= active_entry_price + 5.0) and (live_pe_ltp <= p_open_15m or peak_price - live_pe_ltp >= 6.0) and not (curr_mfi14_pe > prev_mfi14_pe)
                    is_breakout_mb_rejection_pe = is_breakout_trade_pe and (peak_price >= grid.pe_leg.ltp - 5.0 and peak_price <= grid.pe_leg.ltp + 8.0) and (curr_mfi14_pe < prev_mfi14_pe or curr_mfi5_pe >= 80.0 or curr_mfi14_pe >= 68.0) and (live_pe_ltp <= peak_price - 3.0)
                    is_breakout_mfi100_fall_pe = is_breakout_trade_pe and (curr_mfi5_pe >= 100.0 and curr_mfi14_pe < prev_mfi14_pe)
                    is_breakout_3m_fall_pe = is_breakout_trade_pe and (curr_mfi5_pe >= 98.0) and (prev_mfi5_pe - curr_mfi5_pe > 1.0)

                    # Rule #5 Specific Exit: For Recovery Re-Entry PE - Mandatory exit on 15m dual MFI fall
                    is_recovery_mfi_fall_pe = ("Recovery Re-Entry" in entry_type_str_pe if 'entry_type_str_pe' in locals() else False) and (curr_mfi5_pe < prev_mfi5_pe and curr_mfi14_pe < prev_mfi14_pe)

                    # Rule #6 Specific Exit: For Post-Breakdown Oversold Bounce PE - Exit on dual MFI fall or MFI14 fall with price rejection
                    is_post_breakdown_rejection_mfi_fall_pe = ("Post-Breakdown" in entry_type_str_pe if 'entry_type_str_pe' in locals() else False) and ((curr_mfi5_pe < prev_mfi5_pe and curr_mfi14_pe < prev_mfi14_pe) or ((curr_mfi14_pe < prev_mfi14_pe) and (live_pe_ltp < p_open_15m or live_pe_ltp <= peak_price - 3.0)))

                    if ("Post-Breakdown" in entry_type_str_pe if 'entry_type_str_pe' in locals() else False) and (curr_mfi5_pe > prev_mfi5_pe and curr_mfi14_pe > prev_mfi14_pe):
                        active_sl = max(active_sl, active_entry_price - 20.0)

                    # Rule: Exit on same candle closing if 15 min any MFI was falling at entry time and price crossed Upper BB with MFI falling
                    res_bb_pe_check = get_3m_bollinger_bands(smart_api, "BFO", active_contract.symbol_token)
                    ub_price_pe = res_bb_pe_check[1] if (res_bb_pe_check and res_bb_pe_check[1]) else grid.pe_leg.target_epm
                    is_ub_crossed_pe = (peak_price >= ub_price_pe) or (live_pe_ltp >= ub_price_pe)
                    is_mfi_falling_now_pe = (curr_mfi5_pe < prev_mfi5_pe) or (curr_mfi14_pe < prev_mfi14_pe)
                    is_same_candle_ub_mfi_fall_exit_pe = entry_mfi_falling_15m and is_ub_crossed_pe and is_mfi_falling_now_pe

                    # Swing Low Breakout / Retest Strategy Specific Exit Rule PE:
                    # User Rule: Wait after entry if MFI(14) is rising till Upper Bollinger band price rejection or MFI down near or above Upper Band
                    is_swing_trade_pe = ("Swing Low" in entry_type_str_pe if 'entry_type_str_pe' in locals() else False) or (active_strategy_name == "Dynamic Swing Low First Breakout Retest Entry")
                    is_swing_mfi14_rising_pe = (curr_mfi14_pe > prev_mfi14_pe) or (curr_mfi14_pe >= prev_mfi14_pe and curr_mfi5_pe > prev_mfi5_pe)
                    is_swing_ub_rejection_pe = (live_pe_ltp >= ub_price_pe - 2.0 or peak_price >= ub_price_pe - 2.0) and (live_pe_ltp <= p_open_15m or peak_price - live_pe_ltp >= 5.0)
                    is_swing_mfi_down_near_ub_pe = (peak_price >= ub_price_pe - 5.0 or live_pe_ltp >= ub_price_pe - 5.0 or curr_mfi14_pe >= 65.0) and (curr_mfi14_pe < prev_mfi14_pe or curr_mfi5_pe < prev_mfi5_pe)
                    is_swing_exit_pe = is_swing_trade_pe and ((is_swing_mfi14_rising_pe and (is_swing_ub_rejection_pe or is_swing_mfi_down_near_ub_pe) and (live_pe_ltp >= active_entry_price + 8.0 or peak_price >= active_entry_price + 15.0)) or (not is_swing_mfi14_rising_pe and (curr_mfi5_pe < prev_mfi5_pe and curr_mfi14_pe < prev_mfi14_pe)))

                    # Extreme Overbought High Rejection Exit Rule PE:
                    is_overbought_high_rejection_pe = (curr_mfi14_pe >= 70.0 or curr_mfi5_pe >= 80.0) and (live_pe_ltp >= previous_pe_high or peak_price >= previous_pe_high) and (live_pe_ltp < previous_pe_high or live_pe_ltp <= p_open_15m or peak_price - live_pe_ltp >= 5.0)

                    # Profit Booking Exit Logic PE (80+ points OR MFI(14) or both MFI falling in 15m/3m frame)
                    points_gained_pe = live_pe_ltp - active_entry_price
                    is_80pt_profit_booking_pe = (points_gained_pe >= 80.0)
                    mfi_falling_15m_pe = (curr_mfi14_pe < prev_mfi14_pe) or (curr_mfi5_pe < prev_mfi5_pe and curr_mfi14_pe < prev_mfi14_pe)
                    mfi_falling_3m_pe = (mfi14_3m_pe < prev_mfi14_3m_pe) or (mfi5_3m_pe < prev_mfi5_3m_pe and mfi14_3m_pe < prev_mfi14_3m_pe) if ('mfi14_3m_pe' in locals() and 'prev_mfi14_3m_pe' in locals()) else False
                    is_mfi_falling_profit_booking_pe = (points_gained_pe > 0.0) and (mfi_falling_15m_pe or mfi_falling_3m_pe)
                    is_profit_booking_exit_pe = is_80pt_profit_booking_pe or is_mfi_falling_profit_booking_pe

                    # Same Day EOD Mandatory Exit PE (3:25 PM / 3:30 PM cutoff)
                    now_time_str_exit_pe = datetime.now(IST).strftime("%H:%M")
                    is_eod_exit_live_pe = (now_time_str_exit_pe >= "15:25")

                    is_mfi_exit_triggered_pe = is_profit_booking_exit_pe or is_eod_exit_live_pe or is_swing_exit_pe or is_overbought_high_rejection_pe or is_breakout_mfi14_or_both_fall_pe or is_breakout_weak_close_retrace_pe or is_breakout_mb_rejection_pe or is_mb_rejection_pe or is_dual_mfi_falling_pe or is_overbought_mfi14_fall_pe or is_reentry_ub_cross_fall_pe or is_breakout_mfi100_fall_pe or is_breakout_3m_fall_pe or is_recovery_mfi_fall_pe or is_post_breakdown_rejection_mfi_fall_pe or is_same_candle_ub_mfi_fall_exit_pe
                    
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
                        if exit_state_str == "EXIT_SL":
                            pe_sl_hit_today = True
                        
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
                        active_strategy_name = ""
                        initial_hard_sl = 0.0
                        handover_history = []
                        trades_completed += 1
                        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                        # If Hard SL hit (not TSL), allow One-Time Post-SL Recovery Re-Entry
                        if not trailing_active and not recovery_reentry_done:
                            recovery_reentry_eligible = True
                        trailing_active = False
                        peak_price = 0.0
                        recent_ce_low = live_ce_ltp
                        recent_pe_low = live_pe_ltp

                    # 3. Standard Practical Target exit or MFI / 3m Upper BB Target Exit
                    if bot_state != "PE_LONG" or active_contract is None:
                        continue
                    p_mfi_val, p_mfi_prev_val, _ = get_15m_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    p_1h_mfi, p_prev_1h_mfi = get_1h_mfi(smart_api, "BFO", active_contract.symbol_token, period=5)
                    res_bb_pe = get_3m_bollinger_bands(smart_api, "BFO", active_contract.symbol_token)
                    ub_3m_pe = res_bb_pe[1] if res_bb_pe else None
                    
                    # Fetch 3m MFIs (5 and 14) for 3m Upper BB momentum evaluation
                    mfis_3m_pe, prev_mfis_3m_pe = get_mfi_multi_period(smart_api, "BFO", active_contract.symbol_token, "THREE_MINUTE", [5, 14])
                    curr_mfi5_3m_pe = mfis_3m_pe.get(5, 50.0)
                    prev_mfi5_3m_pe = prev_mfis_3m_pe.get(5, 50.0)
                    curr_mfi14_3m_pe = mfis_3m_pe.get(14, 50.0)
                    prev_mfi14_3m_pe = prev_mfis_3m_pe.get(14, 50.0)

                    is_1h_mfi_falling_pe = (p_1h_mfi < p_prev_1h_mfi)
                    is_3m_ub_near_pe = (ub_3m_pe is not None) and (live_pe_ltp >= ub_3m_pe - 5.0)
                    is_40pt_gain_pe = (favorable_gain_pe >= 40.0)

                    # Both 3m MFIs increasing (by at least 1.0 point)
                    both_mfi_increasing_3m_pe = (curr_mfi5_3m_pe >= prev_mfi5_3m_pe + 1.0) and (curr_mfi14_3m_pe >= prev_mfi14_3m_pe + 1.0)
                    # Hold while MFI(14) is rising after MFI(5) reaches 100
                    mfi14_rising_after_100_pe = (curr_mfi5_3m_pe >= 99.0 or p_mfi_val >= 99.0) and (curr_mfi14_3m_pe >= prev_mfi14_3m_pe)
                    hold_due_to_surging_mfi_pe = both_mfi_increasing_3m_pe or mfi14_rising_after_100_pe

                    # Exit allowed if MFI(14) is falling OR both 3m MFIs are falling
                    is_mfi14_falling_3m_pe = (curr_mfi14_3m_pe < prev_mfi14_3m_pe)
                    both_mfi_falling_3m_pe = (curr_mfi5_3m_pe < prev_mfi5_3m_pe) and is_mfi14_falling_3m_pe
                    exit_mfi_confirmed_pe = is_mfi14_falling_3m_pe or both_mfi_falling_3m_pe

                    # Higher Time Frame MFI Hierarchy: follow 3m Upper BB / +40pt profit booking ONLY IF any HTF MFI is falling
                    is_htf_mfi_falling_pe = is_1h_mfi_falling_pe or (p_mfi_val < p_mfi_prev_val)

                    # Book profit at 3m Upper BB or +40pt gain ONLY IF exit is confirmed by falling 3m MFI AND any HTF MFI is falling
                    is_3m_bb_profit_booking_pe = (is_3m_ub_near_pe or is_40pt_gain_pe) and exit_mfi_confirmed_pe and is_htf_mfi_falling_pe and not hold_due_to_surging_mfi_pe
                    
                    is_mfi_tp_hit_pe = (p_mfi_val >= 99.0 and not (curr_mfi14_3m_pe >= prev_mfi14_3m_pe)) or (live_pe_ltp >= active_entry_price + 100.0 and p_mfi_val < p_mfi_prev_val) or is_3m_bb_profit_booking_pe
                    
                    if not trailing_active and (live_pe_ltp >= active_target or is_mfi_tp_hit_pe):
                        tp_reason_pe = "3m Upper BB / +40pt Profit Booking (HTF MFI Falling)" if is_3m_bb_profit_booking_pe else ("MFI Target (100 / 100pt+ & Declining)" if is_mfi_tp_hit_pe else f"Practical Target (₹{active_target:.2f})")
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
                        active_strategy_name = ""
                        initial_hard_sl = 0.0
                        handover_history = []
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

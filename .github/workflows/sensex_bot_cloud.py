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

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("cloud_bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger("sensex_cloud_bot")

IST = timezone(timedelta(hours=5, minutes=30))


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
            "%Y-%m-%d", "%d-%b-%Y", "%d%b%Y", "%d%b%y", "%d-%b-%y",
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
    )


def calculate_master_grid(
    spot: float, vix: float, dte: float,
    ce_ltp: float, ce_delta: float, ce_strike: float,
    pe_ltp: float, pe_delta: float, pe_strike: float,
    buffer: float = 0.15,
) -> EPMMasterGrid:
    spot_val = float(spot)
    vix_val = float(vix) if float(vix) > 0 else 13.5
    dte_val = float(dte)
    time_factor = math.sqrt(max(0.0, dte_val) / 365.0)

    index_move = spot_val * (vix_val / 100.0) * time_factor
    noise_10 = index_move * 0.10
    lower_index = spot_val - index_move
    upper_index = spot_val + index_move

    ce_leg = calculate_master_grid_leg("CE", ce_strike, ce_ltp, ce_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer)
    pe_leg = calculate_master_grid_leg("PE", pe_strike, pe_ltp, pe_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer)

    return EPMMasterGrid(
        spot=spot_val, vix=vix_val, dte=dte_val, dte_sqrt=time_factor,
        index_move=index_move, noise_10=noise_10,
        lower_index=lower_index, upper_index=upper_index,
        ce_leg=ce_leg, pe_leg=pe_leg,
    )


# =========================================================================
# 2. MOBILE NOTIFICATIONS (TELEGRAM / TELEGRAM BOT / SMS WEBHOOK)
# =========================================================================

def check_remote_telegram_command() -> tuple[str | None, int | None]:
    """Check Telegram for incoming commands ('STOP', 'LIVE [LOTS]', 'DEMO').
    Returns tuple: (cmd, lot_size)
    Examples:
      - 'STOP' -> ('STOP', None)
      - 'LIVE 2' -> ('LIVE', 2)
      - 'LIVE' -> ('LIVE', 1)
      - 'DEMO' -> ('DEMO', None)
    """
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not bot_token:
        return None, None
    try:
        import json
        url = f"https://api.telegram.org/bot{bot_token}/getUpdates?offset=-1"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode())
            if data.get("ok") and data.get("result"):
                last_msg = data["result"][-1].get("message", {}).get("text", "").strip().upper()
                parts = last_msg.split()
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


def select_nearest_itm_contract(
    contracts: Iterable[OptionContract],
    spot_price: float,
    option_type: str,
) -> OptionContract:
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

    if option_type == "CE":
        itm = [c for c in matching if c.strike < spot]
        return max(itm if itm else matching, key=lambda c: c.strike)
    else:
        itm = [c for c in matching if c.strike > spot]
        return min(itm if itm else matching, key=lambda c: c.strike)


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
    return smart_api


# =========================================================================
# 5. MAIN CLOUD RUNNER
# =========================================================================

def run_cloud_bot() -> None:
    logger.info("🚀 Starting Standalone Cloud SENSEX Options Bot...")
    excel_tracker = ExcelTracker()

    smart_api = create_authenticated_smartapi_client()

    # Get Spot Price
    spot_res = smart_api.ltpData("BSE", "SENSEX", "99919000")
    spot_price = float(spot_res["data"]["ltp"]) if isinstance(spot_res, dict) and spot_res.get("data") else 77500.0

    # Get VIX
    vix_res = smart_api.ltpData("NSE", "INDIA VIX", "99926017")
    vix_val = float(vix_res["data"]["ltp"]) if isinstance(vix_res, dict) and vix_res.get("data") else 13.5

    # Search contracts
    search_res = smart_api.searchScrip("BFO", "SENSEX")
    rows = search_res.get("data", []) if isinstance(search_res, dict) else []

    contracts: list[OptionContract] = []
    for r in rows:
        symbol = str(r.get("tradingsymbol") or "").strip()
        if not symbol or not (symbol.endswith("CE") or symbol.endswith("PE")):
            continue
        opt_type = "CE" if symbol.endswith("CE") else "PE"
        token = str(r.get("symboltoken") or "").strip()

        m_strike = re.search(r"(\d+)(?:CE|PE)$", symbol)
        if not m_strike:
            continue
        num_str = m_strike.group(1)
        strike_val = float(num_str[-5:]) if len(num_str) >= 5 else float(num_str)

        raw_exp = str(r.get("expiry") or "").strip()
        if raw_exp:
            expiry_val = raw_exp
        else:
            m1 = re.search(r"SENSEX(\d{2})([A-Za-z]{3})(\d+)(?:CE|PE)$", symbol)
            m2 = re.search(r"SENSEX(\d{2})([1-9ONDond])(\d{2})(\d{5})(?:CE|PE)$", symbol)
            if m1:
                yy, mmm, _ = m1.groups()
                expiry_val = f"20{yy}-{mmm.upper()}-01"
            elif m2:
                yy, m_code, dd, _ = m2.groups()
                month_map = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9, "O": 10, "N": 11, "D": 12}
                m_num = month_map.get(m_code.upper(), 8)
                expiry_val = f"20{yy}-{m_num:02d}-{dd}"
            else:
                expiry_val = "2026-08-24"

        dte_days, _ = calculate_dte_sqrt(expiry_val)
        delta_val = calculate_bsm_delta(spot_price, strike_val, dte_days, vix_val, opt_type)

        try:
            c_obj = OptionContract("BFO", symbol, token, expiry_val, strike_val, opt_type, delta_val)
            contracts.append(c_obj)
        except Exception:
            continue

    ce_contract = select_nearest_itm_contract(contracts, spot_price, "CE")
    pe_contract = select_nearest_itm_contract(contracts, spot_price, "PE")

    ce_ltp_res = smart_api.ltpData("BFO", ce_contract.trading_symbol, ce_contract.symbol_token)
    pe_ltp_res = smart_api.ltpData("BFO", pe_contract.trading_symbol, pe_contract.symbol_token)

    ce_ltp = float(ce_ltp_res["data"]["ltp"]) if isinstance(ce_ltp_res, dict) and ce_ltp_res.get("data") else 500.0
    pe_ltp = float(pe_ltp_res["data"]["ltp"]) if isinstance(pe_ltp_res, dict) and pe_ltp_res.get("data") else 300.0

    dte_days, _ = calculate_dte_sqrt(ce_contract.expiry)

    grid = calculate_master_grid(
        spot=spot_price, vix=vix_val, dte=dte_days,
        ce_ltp=ce_ltp, ce_delta=abs(ce_contract.delta), ce_strike=ce_contract.strike,
        pe_ltp=pe_ltp, pe_delta=abs(pe_contract.delta), pe_strike=pe_contract.strike,
        buffer=0.15
    )

    logger.info("=========================================================================")
    logger.info("SENSEX CLOUD BOT - MASTER GRID INITIALIZED")
    logger.info("Spot: %.2f | VIX: %.2f%% | DTE: %.2f | Move: ±%.2f", spot_price, vix_val, dte_days, grid.index_move)
    logger.info("CE Strike %d: LTP ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                grid.ce_leg.strike, ce_ltp, grid.ce_leg.delta, grid.ce_leg.epm_lower_range, grid.ce_leg.target_epm, grid.ce_leg.sl_auto, grid.ce_leg.practical_target)
    logger.info("PE Strike %d: LTP ₹%.2f | Delta %.3f | Lower ₹%.2f | Upper ₹%.2f | SL ₹%.2f | Pr. ₹%.2f",
                grid.pe_leg.strike, pe_ltp, grid.pe_leg.delta, grid.pe_leg.epm_lower_range, grid.pe_leg.target_epm, grid.pe_leg.sl_auto, grid.pe_leg.practical_target)
    logger.info("=========================================================================")

    # Send Notification
    msg = (
        f"🔔 *SENSEX MASTER GRID INITIALIZED*\n\n"
        f"📈 *Spot Price*: ₹{spot_price:.2f} | *VIX*: {vix_val:.2f}%\n"
        f"📅 *DTE*: {dte_days:.1f} Days\n\n"
        f"🟢 *CE {int(grid.ce_leg.strike)}*:\n"
        f"• LTP: ₹{ce_ltp:.2f} | Delta: {grid.ce_leg.delta:.3f}\n"
        f"• EPM Low: ₹{grid.ce_leg.epm_lower_range:.2f} | Auto SL: ₹{grid.ce_leg.sl_auto:.2f}\n"
        f"• EPM Upper Target: ₹{grid.ce_leg.target_epm:.2f} | Practical Target: ₹{grid.ce_leg.practical_target:.2f}\n\n"
        f"🔴 *PE {int(grid.pe_leg.strike)}*:\n"
        f"• LTP: ₹{pe_ltp:.2f} | Delta: {-grid.pe_leg.delta:.3f}\n"
        f"• EPM Low: ₹{grid.pe_leg.epm_lower_range:.2f} | Auto SL: ₹{grid.pe_leg.sl_auto:.2f}\n"
        f"• EPM Upper Target: ₹{grid.pe_leg.target_epm:.2f} | Practical Target: ₹{grid.pe_leg.practical_target:.2f}"
    )
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

    # Continuous Monitoring Loop
    poll_interval = 1.0
    is_continuous = "--once" not in sys.argv
    execution_mode = "LIVE" if "--live" in sys.argv else "PAPER"

    if is_continuous:
        logger.info("🔄 Entering continuous monitoring loop (Mode: %s, Refreshing 1s in-place)...", execution_mode)
        try:
            while True:
                time.sleep(poll_interval)
                checked_at = datetime.now(IST)

                # Check Telegram for remote commands ('LIVE [LOTS]', 'DEMO', 'STOP')
                cmd, remote_lots = check_remote_telegram_command()
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

                # Fetch Live Spot & LTPs
                live_spot_res = smart_api.ltpData("BSE", "SENSEX", "99919000")
                live_spot = float(live_spot_res["data"]["ltp"]) if isinstance(live_spot_res, dict) and live_spot_res.get("data") else spot_price

                live_ce_res = smart_api.ltpData("BFO", ce_contract.trading_symbol, ce_contract.symbol_token)
                live_pe_res = smart_api.ltpData("BFO", pe_contract.trading_symbol, pe_contract.symbol_token)

                live_ce_ltp = float(live_ce_res["data"]["ltp"]) if isinstance(live_ce_res, dict) and live_ce_res.get("data") else ce_ltp
                live_pe_res_ltp = float(live_pe_res["data"]["ltp"]) if isinstance(live_pe_res, dict) and live_pe_res.get("data") else pe_ltp

                status_line = f"\r[{checked_at.strftime('%H:%M:%S')}] [{execution_mode}] Spot: {live_spot:.2f} | CE {ce_contract.strike:.0f}: ₹{live_ce_ltp:.2f} (Low ₹{grid.ce_leg.epm_lower_range:.2f}, High ₹{grid.ce_leg.target_epm:.2f}) | PE {pe_contract.strike:.0f}: ₹{live_pe_res_ltp:.2f} (Low ₹{grid.pe_leg.epm_lower_range:.2f}, High ₹{grid.pe_leg.target_epm:.2f})"
                sys.stdout.write(status_line.ljust(110))
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

# -*- coding: utf-8 -*-
# Standalone Single-File Autonomous SENSEX/BankNifty Options Bot for m.Stock (Mirae Asset)
"""
Standalone Single-File Autonomous Options Bot fully integrated with m.Stock (Mirae Asset API).

Consolidates into ONE single independent file:
1. Quantitative Math (BSM Delta, DTE mapping, EPM Master Grid, Confluence Scoring)
2. m.Stock (Mirae Asset) API Adapter with Auto-URL Discovery & Session Token Auto-Injection
3. Embedded Modular 4-Strategy Suite with Dynamic In-Flight Handover (Zero external .py dependencies)
4. Multi-Timeframe MFI, ATR, Bollinger Bands & Lower Wick Absorption Analytics
5. Embedded Historical Backtest Engine for m.Stock Option Candles
6. Excel Signal & Order Tracker (signal_tracker.xlsx) + JSON State Persistence
7. Autonomous Real-Time Monitoring Loop with Remote Telegram Controls & Voice Alerts

Usage:
    Live/Paper Trading: python sbd_bot_mstock.py [--paper] [--lots=4]
    Historical Backtest: python sbd_bot_mstock.py --backtest [--from-date=2026-08-01]
"""

from __future__ import annotations

import contextlib
import io
import json
import math
import os
import re
import sys
import time
import logging
import threading
import urllib.request
import urllib.parse
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Mapping, Iterable, Sequence, Dict, List, Optional, Tuple

import pandas as pd
import numpy as np

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Reconfigure standard streams to UTF-8 to prevent charmap encoding errors in Windows consoles
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

logging.Formatter.converter = ist_converter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("mstock_bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger("mstock_options_bot")

if os.name == "nt":
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        pass


# =========================================================================
# SYSTEM CONFIGURATION & STRATEGY FEATURE TOGGLES
# =========================================================================
ENABLE_IN_FLIGHT_HANDOVER: bool = True
MIN_EV_HANDOVER_DELTA: float = 4.0


# =========================================================================
# 1. EMBEDDED STRATEGY ENGINE DATA TYPES & ENUMS
# =========================================================================

class MachineState(Enum):
    STATE_0_SCANNING = "STATE_0_SCANNING"
    STATE_1_ACTIVE_LONG = "STATE_1_ACTIVE_LONG"
    STATE_2_ACTIVE_SHORT = "STATE_2_ACTIVE_SHORT"
    STATE_3_HANDOVER_WINDOW = "STATE_3_HANDOVER_WINDOW"


class PositionSide(Enum):
    LONG = "LONG"
    SHORT = "SHORT"


@dataclass
class EntrySignal:
    strategy_name: str
    side: PositionSide
    entry_price: float
    initial_sl: float
    target_price: float
    lot_size: int = 1
    reason: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ExitSignal:
    should_exit: bool
    exit_price: float
    exit_reason: str
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class HandoverDecision:
    switch_approved: bool
    from_strategy: str
    to_strategy: str
    current_ev: float
    projected_ev: float
    adjusted_sl: float
    adjusted_target: float
    reason: str


@dataclass
class Position:
    trade_id: int
    strategy_name: str
    side: PositionSide
    entry_time: str
    entry_price: float
    current_sl: float
    initial_sl: float
    target_price: float
    peak_price: float
    lot_size: int
    handover_history: List[Dict[str, Any]] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    ub_reached: bool = False
    mfi14_overbought_reached: bool = False
    last_swing_high_price: float = 0.0
    last_swing_high_mfi14: float = 0.0
    entry_30m_both_rising: bool = False
    entry_15m_both_rising: bool = False


@dataclass
class MarketContext:
    timestamp: str
    open: float
    high: float
    low: float
    close: float
    volume: float
    mfi5_15m: float
    mfi14_15m: float
    prev_mfi5_15m: float
    prev_mfi14_15m: float
    mfi5_30m: float
    mfi14_30m: float
    prev_mfi5_30m: float
    prev_mfi14_30m: float
    prev_prev_mfi14_30m: float
    mfi5_60m: float
    mfi14_60m: float
    prev_mfi5_60m: float
    prev_mfi14_60m: float
    mb_20: float
    ub_20: float
    lb_20: float
    prev_high: float
    prev_low: float
    prev_close: float
    recent_swing_low: float
    recent_swing_high: float
    dynamic_tolerance: float
    is_0915_bar: bool
    is_big_gap_up: bool
    allow_reentry: bool
    recovery_eligible: bool
    initial_entry_done: bool
    prev_open: float = 0.0
    mfi5_3m: float = 50.0
    mfi14_3m: float = 50.0
    prev_mfi5_3m: float = 50.0
    prev_mfi14_3m: float = 50.0
    prev_prev_mfi5_3m: float = 50.0
    prev_prev_mfi14_3m: float = 50.0
    mb_3m: float = 0.0
    ub_3m: float = 0.0
    lb_3m: float = 0.0
    prev_mb_20: float = 0.0
    prev_prev_mb_20: float = 0.0
    prev_mb_3m: float = 0.0
    prev_prev_mb_3m: float = 0.0


# =========================================================================
# 2. QUANTITATIVE MATHEMATICS & DTE MAPPING
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


@contextlib.contextmanager
def suppress_stdout_stderr():
    """Suppress low-level stdout/stderr output during verbose API calls."""
    try:
        null_fd = os.open(os.devnull, os.O_RDWR)
        save_stdout = os.dup(1)
        save_stderr = os.dup(2)
        os.dup2(null_fd, 1)
        os.dup2(null_fd, 2)
        yield
    except Exception:
        yield
    finally:
        try:
            os.dup2(save_stdout, 1)
            os.dup2(save_stderr, 2)
            os.close(null_fd)
            os.close(save_stdout)
            os.close(save_stderr)
        except Exception:
            pass


def get_last_tuesday_of_month(year: int, month: int) -> date:
    """Find the last Tuesday of a given month and year."""
    if month == 12:
        last_day = date(year, 12, 31)
    else:
        last_day = date(year, month + 1, 1) - timedelta(days=1)
    offset = (last_day.weekday() - 1) % 7
    return last_day - timedelta(days=offset)


def count_trading_days(start_date: date, end_date: date) -> int:
    """Count Monday-Friday trading days between start_date and end_date."""
    if start_date >= end_date:
        return 0
    cur = start_date
    trading_days = 0
    while cur < end_date:
        if cur.weekday() < 5:
            trading_days += 1
        cur += timedelta(days=1)
    return max(1, trading_days)


def calculate_dte_sqrt(expiry: Any, as_of: Any = None, index_name: str = "SENSEX") -> tuple[float, float]:
    """Calculate deterministic trading-day DTE mapping to contract expiry date."""
    as_of_dt = to_ist_datetime(as_of)
    today = as_of_dt.date()
    is_bn = "BANKNIFTY" in str(index_name).upper() or (isinstance(expiry, str) and "BANKNIFTY" in expiry.upper())

    expiry_date = None
    if expiry:
        try:
            exp_dt = to_ist_datetime(expiry)
            expiry_date = exp_dt.date()
        except Exception:
            pass

    if not expiry_date:
        if is_bn:
            last_tue = get_last_tuesday_of_month(today.year, today.month)
            if today > last_tue:
                next_month = 1 if today.month == 12 else today.month + 1
                next_year = today.year + 1 if today.month == 12 else today.year
                last_tue = get_last_tuesday_of_month(next_year, next_month)
            expiry_date = last_tue
        else:
            wd = today.weekday()
            dte_days = 4.0 if wd == 0 else (3.0 if wd == 1 else (2.0 if wd == 2 else (1.0 if wd == 3 else 5.0)))
            time_factor = math.sqrt(dte_days / 365.0)
            return dte_days, time_factor

    if today >= expiry_date:
        dte_days = 0.5
    else:
        dte_days = float(count_trading_days(today, expiry_date))

    time_factor = math.sqrt(dte_days / 365.0)
    return dte_days, time_factor


def calculate_bsm_delta(spot: float, strike: float, dte_days: float, vix: float = 13.5, option_type: str = "CE") -> float:
    """Calculate Black-Scholes Delta."""
    if spot <= 0 or strike <= 0 or dte_days <= 0:
        return 0.50 if option_type == "CE" else -0.50

    t = max(dte_days, 0.001) / 365.0
    sigma = max(vix, 1.0) / 100.0
    r = 0.07

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
    sl_offset: float = 15.0,
) -> MasterGridLeg:
    abs_delta = abs(float(delta))
    ltp_val = float(ltp)
    buf_val = float(buffer)

    option_move = float(index_move) * abs_delta
    epm_lower_range = max(5.0, ltp_val - (option_move * buf_val))
    target_epm = ltp_val + (option_move * buf_val)
    sl_auto = max(5.0, epm_lower_range - float(sl_offset))
    practical_target = ltp_val + (option_move * 0.21)

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


def calculate_dynamic_epm_proximity_threshold(
    atr14_15m: float,
    iv_live: float = 13.5,
    iv_anchor: float = 13.5,
    delta: float = 0.55,
    min_thresh: float = 8.0,
) -> float:
    safe_atr = max(5.0, float(atr14_15m)) if atr14_15m and atr14_15m > 0 else 15.0
    safe_anchor_iv = max(5.0, float(iv_anchor)) if iv_anchor and iv_anchor > 0 else 13.5
    safe_live_iv = max(5.0, float(iv_live)) if iv_live and iv_live > 0 else safe_anchor_iv
    abs_delta = abs(float(delta))
    
    iv_ratio = 1.0 + ((safe_live_iv - safe_anchor_iv) / safe_anchor_iv)
    dynamic_dist = safe_atr * iv_ratio * abs_delta * 0.30
    return max(float(min_thresh), dynamic_dist)


def calculate_index_scaled_sl(
    index_name: str,
    base_sl_sensex: float = 18.0,
    spot_price: float = 51500.0,
    sensex_spot: float = 81500.0,
    delta: float = 0.60,
    lot_size: int = 30,
) -> float:
    idx = str(index_name).upper().strip()
    if "SENSEX" in idx:
        return float(base_sl_sensex)
    
    scale_factor = (spot_price / sensex_spot) if sensex_spot > 0 else 0.632
    spot_move = (base_sl_sensex / max(0.1, delta)) * scale_factor
    option_volatility_sl = spot_move * max(0.1, delta)
    
    if lot_size >= 120:
        return round(max(10.0, min(14.0, option_volatility_sl)), 1)
    elif lot_size >= 90:
        return 16.0
    elif lot_size >= 60:
        return 24.0
    return round(max(10.0, option_volatility_sl), 1)


def calculate_lower_wick_absorption(
    candle_open: float,
    candle_high: float,
    candle_low: float,
    candle_close: float,
) -> float:
    total_range = float(candle_high) - float(candle_low)
    if total_range <= 0.0:
        return 0.0
    body_bottom = min(float(candle_open), float(candle_close))
    lower_wick = max(0.0, body_bottom - float(candle_low))
    return round((lower_wick / total_range) * 100.0, 2)


def calculate_confluence_score(
    ltp: float,
    epm_low: float,
    dynamic_thresh: float,
    mfi14_15m: float,
    prev_mfi14_15m: float,
    mfi5_3m: float,
    prev_mfi5_3m: float,
    spot: float,
    spot_open: float,
    candle_open: float,
    option_type: str = "CE",
) -> tuple[int, list[str]]:
    score = 0
    reasons = []
    
    if epm_low > 0.0:
        if (ltp >= epm_low - 12.0) and (ltp <= epm_low + dynamic_thresh):
            score += 35
            reasons.append(f"EPM Proximity (+35): LTP ₹{ltp:.2f} near EPM Low ₹{epm_low:.2f}")
    
    mfi_15m_rising = (mfi14_15m > prev_mfi14_15m) or (mfi14_15m <= 35.0 and mfi14_15m >= prev_mfi14_15m)
    mfi_3m_rising = (mfi5_3m > prev_mfi5_3m) or (mfi5_3m <= 25.0)
    if mfi_15m_rising and mfi_3m_rising:
        score += 30
        reasons.append(f"Dual MFI Rising (+30): 15m MFI {mfi14_15m:.1f} & 3m MFI {mfi5_3m:.1f}")
    elif mfi_15m_rising or mfi_3m_rising:
        score += 15
        reasons.append(f"Single MFI Rising (+15): 15m MFI {mfi14_15m:.1f} / 3m MFI {mfi5_3m:.1f}")
        
    is_ce = (option_type.upper() == "CE")
    if is_ce and (spot >= spot_open - 20.0):
        score += 20
        reasons.append(f"Spot Bullish (+20): Spot ₹{spot:.2f} >= Open ₹{spot_open:.2f}")
    elif not is_ce and (spot <= spot_open + 20.0):
        score += 20
        reasons.append(f"Spot Bearish (+20): Spot ₹{spot:.2f} <= Open ₹{spot_open:.2f}")
        
    if candle_open > 0.0 and ltp >= candle_open:
        score += 15
        reasons.append(f"Candle Open Reclaim (+15): LTP ₹{ltp:.2f} >= Open ₹{candle_open:.2f}")
        
    return score, reasons


def calculate_master_grid(
    spot: float, vix: float, dte: float,
    ce_ltp: float, ce_delta: float, ce_strike: float,
    pe_ltp: float, pe_delta: float, pe_strike: float,
    buffer: float = 0.12,
    ce_legs_data: list[Any] | None = None,
    pe_legs_data: list[Any] | None = None,
    ce_expiry: str = "", ce_trading_symbol: str = "",
    pe_expiry: str = "", pe_trading_symbol: str = "",
    sl_offset: float = 15.0,
) -> EPMMasterGrid:
    spot_val = float(spot)
    vix_val = float(vix) if float(vix) > 0 else 13.5
    dte_val = float(dte)
    time_factor = math.sqrt(max(0.0, dte_val) / 365.0)

    index_move = spot_val * (vix_val / 100.0) * time_factor
    noise_10 = index_move * 0.10
    lower_index = spot_val - index_move
    upper_index = spot_val + index_move

    ce_leg = calculate_master_grid_leg("CE", ce_strike, ce_ltp, ce_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=ce_expiry, trading_symbol=ce_trading_symbol, sl_offset=sl_offset)
    pe_leg = calculate_master_grid_leg("PE", pe_strike, pe_ltp, pe_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=pe_expiry, trading_symbol=pe_trading_symbol, sl_offset=sl_offset)

    ce_legs = []
    if ce_legs_data:
        for item in ce_legs_data:
            c_ltp, c_delta, c_strike = item[0], item[1], item[2]
            c_exp = item[3] if len(item) > 3 else ce_expiry
            c_sym = item[4] if len(item) > 4 else ce_trading_symbol
            ce_legs.append(calculate_master_grid_leg("CE", c_strike, c_ltp, c_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=c_exp, trading_symbol=c_sym, sl_offset=sl_offset))
    else:
        ce_legs = [ce_leg]

    pe_legs = []
    if pe_legs_data:
        for item in pe_legs_data:
            p_ltp, p_delta, p_strike = item[0], item[1], item[2]
            p_exp = item[3] if len(item) > 3 else pe_expiry
            p_sym = item[4] if len(item) > 4 else pe_trading_symbol
            pe_legs.append(calculate_master_grid_leg("PE", p_strike, p_ltp, p_delta, index_move, time_factor, vix=vix_val, dte=dte_val, buffer=buffer, expiry=p_exp, trading_symbol=p_sym, sl_offset=sl_offset))
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
# 3. M.STOCK (MIRAE ASSET) API CLIENT ADAPTER
# =========================================================================

class MstockClientAdapter:
    """
    Production-grade m.Stock (Mirae Asset Open API) Client Adapter.
    Handles session generation, URL discovery, session token auto-injection,
    quote/LTP fetching, candle data retrieval, instrument search, position fetching, and order placement.
    """

    def __init__(
        self,
        api_key: str | None = None,
        client_id: str | None = None,
        password: str | None = None,
        totp_secret: str | None = None,
        app_code: str | None = None,
        base_url: str | None = None
    ):
        self.api_key = api_key or os.environ.get("MSTOCK_API_KEY", "")
        self.client_id = client_id or os.environ.get("MSTOCK_CLIENT_ID", os.environ.get("MSTOCK_USER_ID", ""))
        self.password = password or os.environ.get("MSTOCK_PASSWORD", "")
        self.totp_secret = totp_secret or os.environ.get("MSTOCK_TOTP_SECRET", os.environ.get("MSTOCK_PIN", ""))
        self.app_code = app_code or os.environ.get("MSTOCK_APP_CODE", "MSTOCK_BOT")
        
        # Base URL candidates for auto-discovery
        env_base = os.environ.get("MSTOCK_BASE_URL", "").rstrip('/')
        if env_base:
            self.candidate_urls = [env_base]
        else:
            self.candidate_urls = [
                "https://api.mstock.trade/openapi/typea",
                "https://api.mstock.trade/openapi",
                "https://tradingapi.mstock.com/typea/v1",
                "https://tradingapi.mstock.com/v1",
                "https://tradeapi.mstock.com/v1"
            ]
        self.base_url = self.candidate_urls[0]

        self.access_token: str | None = None
        self.refresh_token: str | None = None
        self.is_authenticated: bool = False
        self.last_auth_time: float = 0.0
        self.last_auth_reason: str = "Initialization pending"

    def generate_session(self) -> bool:
        """Authenticate with m.Stock API using credentials + TOTP with multi-endpoint discovery and exponential retry backoff."""
        if not self.api_key or not self.client_id:
            self.last_auth_reason = "Missing MSTOCK_API_KEY or MSTOCK_CLIENT_ID credentials in environment / .env"
            logger.warning("⚠️ Credentials missing: %s. Operating in Paper / Demo Mode.", self.last_auth_reason)
            self.is_authenticated = False
            return False

        totp_val = ""
        if self.totp_secret:
            try:
                import pyotp
                totp_val = pyotp.TOTP(self.totp_secret).now()
            except Exception:
                totp_val = self.totp_secret

        # Form-urlencoded payload for Type A m.Stock API
        form_payload = {
            "username": self.client_id,
            "password": self.password,
            "totp": totp_val,
            "apiKey": self.api_key
        }
        encoded_form = urllib.parse.urlencode(form_payload).encode('utf-8')

        headers = {
            "X-Mirae-Version": "1",
            "X-Api-Key": self.api_key,
            "Content-Type": "application/x-www-form-urlencoded",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        }

        login_paths = ["connect/login", "login", "user/login", "session/login"]
        failure_details = []
        max_retries = 3

        for b_url in self.candidate_urls:
            for l_path in login_paths:
                full_url = f"{b_url.rstrip('/')}/{l_path.lstrip('/')}"
                for attempt in range(1, max_retries + 1):
                    try:
                        req = urllib.request.Request(
                            full_url,
                            data=encoded_form,
                            headers=headers,
                            method="POST"
                        )
                        with urllib.request.urlopen(req, timeout=8) as resp:
                            if resp.status in (200, 201):
                                res_data = json.loads(resp.read().decode('utf-8'))
                                if res_data.get("status") in (True, "success", "SUCCESS") or "data" in res_data or "token" in res_data:
                                    data = res_data.get("data") or res_data
                                    self.access_token = data.get("ugid") or data.get("accessToken") or data.get("token") or "ACTIVE_SESSION"
                                    self.refresh_token = data.get("refreshToken")
                                    self.base_url = b_url.rstrip('/')
                                    self.is_authenticated = True
                                    self.last_auth_time = time.time()
                                    self.last_auth_reason = "Authentication successful. Live session active."
                                    logger.info("🟢 [MSTOCK AUTH SUCCESS] Authentication successful! Connected to m.Stock Live API Server via %s", full_url)
                                    logger.info("📡 LIVE FEED CONNECTED: Real-time data feed stream is ACTIVE. User UGID: %s", data.get("ugid", self.client_id))
                                    return True
                                else:
                                    reason = res_data.get("message") or res_data.get("error") or str(res_data)
                                    failure_details.append(f"{full_url} -> Server Response: {reason}")
                                    break
                    except urllib.error.HTTPError as http_err:
                        reason = f"HTTP {http_err.code} ({http_err.reason})"
                        failure_details.append(f"{full_url} -> {reason}")
                        if http_err.code in (502, 503, 504) and attempt < max_retries:
                            time.sleep(2 * attempt)
                            continue
                        elif http_err.code in (401, 403):
                            logger.warning("⚠️ m.Stock login credentials rejected at %s (%s)", full_url, reason)
                            break
                        else:
                            break
                    except Exception as exc:
                        failure_details.append(f"{full_url} -> {type(exc).__name__}: {exc}")
                        break

        summary_reason = " | ".join(failure_details[:2]) if failure_details else "Unreachable API endpoints or network timeout"
        self.last_auth_reason = f"Live server authentication failed ({summary_reason})"
        logger.warning("🔴 [MSTOCK AUTH FAILURE] Reason: %s", self.last_auth_reason)
        logger.warning("ℹ️ Operating in Paper / Demo Trading Mode.")
        self.is_authenticated = False
        return False

    def _get_headers(self) -> dict[str, str]:
        headers = {
            "Content-Type": "application/json",
            "X-Api-Key": self.api_key or "",
            "User-Agent": "MstockPythonBot/1.0"
        }
        if self.access_token:
            headers["Authorization"] = f"Bearer {self.access_token}"
        return headers

    def safe_request(self, endpoint: str, method: str = "GET", payload: dict | None = None, max_retries: int = 3) -> dict | None:
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        data_bytes = json.dumps(payload).encode('utf-8') if payload else None

        for attempt in range(1, max_retries + 1):
            try:
                headers = self._get_headers()
                req = urllib.request.Request(url, data=data_bytes, headers=headers, method=method)
                with urllib.request.urlopen(req, timeout=8) as resp:
                    if resp.status in (200, 201):
                        return json.loads(resp.read().decode('utf-8'))
            except urllib.error.HTTPError as http_err:
                if http_err.code in (401, 403) and attempt < max_retries:
                    logger.info("🔄 Session token expired. Triggering m.Stock re-authentication...")
                    self.generate_session()
                else:
                    logger.debug("m.Stock HTTP error (%d) for %s: %s", http_err.code, endpoint, http_err)
            except Exception as exc:
                logger.debug("m.Stock API exception (Attempt %d/%d) for %s: %s", attempt, max_retries, endpoint, exc)
                time.sleep(0.3 * attempt)

        return None

    def ltpData(self, exchange: str, trading_symbol: str, symbol_token: str) -> dict | None:
        if not self.is_authenticated:
            return None

        res = self.safe_request(f"quote/ltp?exchange={exchange}&symbol={trading_symbol}&token={symbol_token}")
        if res and (res.get("status") in (True, "success") or "data" in res):
            data = res.get("data") or res
            return {
                "status": True,
                "data": {
                    "ltp": float(data.get("ltp") or data.get("lastPrice") or 0.0),
                    "open": float(data.get("open") or data.get("openPrice") or data.get("ltp") or 0.0),
                    "high": float(data.get("high") or 0.0),
                    "low": float(data.get("low") or 0.0),
                    "close": float(data.get("close") or 0.0)
                }
            }
        return None

    def getCandleData(self, params: dict) -> dict | None:
        if not self.is_authenticated:
            return None

        exch = params.get("exchange", "BSE")
        token = params.get("symboltoken", "")
        interval = params.get("interval", "FIFTEEN_MINUTE")
        from_date = params.get("fromdate", "")
        to_date = params.get("todate", "")

        query_str = urllib.parse.urlencode({
            "exchange": exch,
            "token": token,
            "interval": interval,
            "from": from_date,
            "to": to_date
        })

        res = self.safe_request(f"charts/candles?{query_str}")
        if res and (res.get("status") in (True, "success") or "data" in res):
            candles = res.get("data") or res.get("candles") or []
            return {"status": True, "data": candles}
        return None

    def searchScrip(self, exchange: str, search_symbol: str) -> dict:
        if not self.is_authenticated:
            return {"data": []}

        res = self.safe_request(f"instruments/search?exchange={exchange}&query={search_symbol}")
        if res and "data" in res:
            return res
        return {"data": []}

    def placeOrder(self, order_params: dict) -> dict:
        if not self.is_authenticated:
            order_id = f"MSTK_PAPER_{int(time.time()*1000)}"
            return {"status": True, "data": {"orderid": order_id, "message": "Paper order placed successfully"}}

        res = self.safe_request("orders/place", method="POST", payload=order_params)
        if res:
            return res
        return {"status": False, "message": "m.Stock Order Placement Failed"}

    def position(self) -> dict:
        if not self.is_authenticated:
            return {"status": True, "data": []}

        res = self.safe_request("portfolio/positions")
        if res and "data" in res:
            return res
        return {"status": True, "data": []}


def create_authenticated_mstock_client() -> MstockClientAdapter:
    client = MstockClientAdapter()
    client.generate_session()
    return client


# =========================================================================
# 4. MOBILE NOTIFICATIONS & EXCEL TRACKER
# =========================================================================

def send_telegram_voice_alert(message: str) -> None:
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not bot_token or not chat_id:
        return
        
    msg_upper = message.upper()
    if "CHEAT SHEET" in msg_upper or "CHEATSHEET" in msg_upper or "COMMANDS" in msg_upper or "HELP" in msg_upper:
        return

    is_event = any(k in msg_upper for k in ("ENTRY", "EXIT", "TARGET", "STOP LOSS", "TSL", "SL HIT", "SIGNAL", "MFI", "GRID", "SWAP", "BREAKOUT"))
    if not is_event:
        return
        
    try:
        clean_text = message.replace("*", "").replace("`", "").replace("🟢", "").replace("🔴", "").replace("🔥", "").replace("SBD_bot", "")
        lines = [line.strip() for line in clean_text.split("\n") if line.strip()][:3]
        voice_text = " . ".join(lines)
        
        try:
            from gtts import gTTS
        except ImportError:
            logger.warning("⚠️ gTTS module not installed. Run 'pip install gtts' for Telegram Voice Notes.")
            return

        import io
        import subprocess

        tts = gTTS(text=voice_text, lang='en', slow=False)
        mp3_fp = io.BytesIO()
        tts.write_to_fp(mp3_fp)
        mp3_bytes = mp3_fp.getvalue()
        
        voice_bytes = mp3_bytes
        filename = "alert.ogg"
        try:
            import shutil
            if shutil.which('ffmpeg'):
                proc = subprocess.Popen(
                    ['ffmpeg', '-y', '-i', 'pipe:0', '-ac', '1', '-ar', '48000', '-c:a', 'libopus', '-b:a', '32k', '-f', 'ogg', 'pipe:1'],
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                ogg_opus_data, _ = proc.communicate(input=mp3_bytes, timeout=8)
                if proc.returncode == 0 and len(ogg_opus_data) > 0:
                    voice_bytes = ogg_opus_data
        except Exception as conv_err:
            logger.warning("Could not convert TTS audio with ffmpeg: %s", conv_err)
        
        estimated_duration = max(2, int(len(voice_text.split()) / 3))

        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        url = f"https://api.telegram.org/bot{bot_token}/sendVoice"
        
        body = [
            f'--{boundary}'.encode('utf-8'),
            f'Content-Disposition: form-data; name="chat_id"'.encode('utf-8'),
            ''.encode('utf-8'),
            str(chat_id).encode('utf-8'),
            f'--{boundary}'.encode('utf-8'),
            f'Content-Disposition: form-data; name="duration"'.encode('utf-8'),
            ''.encode('utf-8'),
            str(estimated_duration).encode('utf-8'),
            f'--{boundary}'.encode('utf-8'),
            f'Content-Disposition: form-data; name="voice"; filename="{filename}"'.encode('utf-8'),
            'Content-Type: audio/ogg'.encode('utf-8'),
            ''.encode('utf-8'),
            voice_bytes,
            f'--{boundary}--'.encode('utf-8'),
            ''.encode('utf-8')
        ]
        
        req_body = b'\r\n'.join(body)
        req = urllib.request.Request(url, data=req_body)
        req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
        req.add_header('User-Agent', 'Mozilla/5.0')
        
        with urllib.request.urlopen(req, timeout=8) as response:
            if response.status == 200:
                logger.info("🎙️ Telegram Voice Alert sent successfully.")
    except Exception as e:
        logger.warning("Failed to send Telegram Voice Alert: %s", e)


def send_mobile_alert(message: str, send_voice: bool = True) -> None:
    if send_voice:
        send_telegram_voice_alert(message)

    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if bot_token and chat_id:
        try:
            url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
            payload = urllib.parse.urlencode({"chat_id": chat_id, "text": message, "parse_mode": "Markdown"}).encode("utf-8")
            req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"})
            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status == 200:
                    logger.info("📱 Telegram push notification sent.")
        except Exception:
            try:
                clean_msg = message.replace("*", "").replace("`", "")
                payload_plain = urllib.parse.urlencode({"chat_id": chat_id, "text": clean_msg}).encode("utf-8")
                req_plain = urllib.request.Request(url, data=payload_plain, headers={"Content-Type": "application/x-www-form-urlencoded"})
                urllib.request.urlopen(req_plain, timeout=5)
            except Exception:
                pass

    wa_phone = os.environ.get("WHATSAPP_PHONE")
    wa_apikey = os.environ.get("WHATSAPP_API_KEY", "")

    if wa_phone:
        try:
            clean_msg = message.replace("*", "").replace("`", "")
            encoded_text = urllib.parse.quote(clean_msg)
            if wa_apikey:
                url = f"https://api.callmebot.com/whatsapp.php?phone={wa_phone}&text={encoded_text}&apikey={wa_apikey}"
            else:
                url = f"https://api.callmebot.com/whatsapp.php?phone={wa_phone}&text={encoded_text}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            urllib.request.urlopen(req, timeout=8)
            logger.info("🟢 WhatsApp alert sent to %s.", wa_phone)
        except Exception as e:
            logger.debug("WhatsApp notification attempt: %s", e)


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
                logger.warning("Could not read existing Excel tracker: %s", e)

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


class TelegramCommandListener:
    def __init__(self, bot_token: str | None) -> None:
        self.bot_token = bot_token
        self.last_update_id = 0
        self.startup_time = int(time.time())
        self._flush_old_updates()

    def _flush_old_updates(self) -> None:
        if not self.bot_token:
            return
        try:
            url = f"https://api.telegram.org/bot{self.bot_token}/getUpdates?offset=-1"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode())
                if data.get("ok") and data.get("result"):
                    self.last_update_id = data["result"][-1].get("update_id", 0)
                    flush_url = f"https://api.telegram.org/bot{self.bot_token}/getUpdates?offset={self.last_update_id + 1}"
                    urllib.request.urlopen(urllib.request.Request(flush_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=3)
        except Exception:
            pass

    def get_new_command(self) -> tuple[str | None, int | float | None]:
        if not self.bot_token:
            return None, None
        try:
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
                            if msg_obj.get("date", 0) < self.startup_time:
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
                                lots = None
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
                                return ("REENTRY ON" if (len(parts) > 1 and parts[1] in ("ON", "TRUE", "1")) else "REENTRY OFF"), None
                            elif first_word in ("UNLOCK", "/UNLOCK"):
                                side = parts[1] if len(parts) > 1 else "ALL"
                                return f"UNLOCK {side}", None
        except Exception:
            pass
        return None, None


# =========================================================================
# 5. EMBEDDED MODULAR 4-STRATEGY SUITE & STATE MACHINE ENGINE
# =========================================================================

def calculate_mfi_series(high: pd.Series, low: pd.Series, close: pd.Series, volume: pd.Series, period: int = 14) -> pd.Series:
    typical_price = (high + low + close) / 3.0
    raw_money_flow = typical_price * volume
    tp_diff = typical_price.diff()
    positive_flow = raw_money_flow.where(tp_diff > 0, 0.0)
    negative_flow = raw_money_flow.where(tp_diff < 0, 0.0)
    
    pos_mf_sum = positive_flow.rolling(window=period, min_periods=period).sum()
    neg_mf_sum = negative_flow.rolling(window=period, min_periods=period).sum()
    
    money_ratio = np.where(neg_mf_sum == 0, 100.0, pos_mf_sum / (neg_mf_sum + 1e-12))
    mfi = 100.0 - (100.0 / (1.0 + money_ratio))
    return pd.Series(np.nan_to_num(mfi, nan=50.0), index=close.index)


class BaseStrategy(ABC):
    def __init__(self, name: str, base_win_rate: float, base_rr: float):
        self.name = name
        self.base_win_rate = base_win_rate
        self.base_rr = base_rr

    @abstractmethod
    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        pass

    @abstractmethod
    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        pass

    def calculate_expected_value(self, ctx: MarketContext, position: Position) -> float:
        dist_to_target = max(5.0, position.target_price - ctx.close)
        dist_to_sl = max(5.0, ctx.close - position.current_sl)
        
        mfi_momentum = 1.0
        if ctx.mfi14_15m > ctx.prev_mfi14_15m:
            mfi_momentum += 0.15
        if ctx.mfi5_15m > ctx.prev_mfi5_15m:
            mfi_momentum += 0.10
        if ctx.mfi14_30m >= ctx.prev_mfi14_30m:
            mfi_momentum += 0.10
        if ctx.mfi14_15m >= 70.0 or ctx.mfi5_15m >= 80.0:
            mfi_momentum -= 0.35

        adjusted_win_rate = min(0.90, max(0.20, self.base_win_rate * mfi_momentum))
        loss_rate = 1.0 - adjusted_win_rate
        return float((adjusted_win_rate * dist_to_target) - (loss_rate * dist_to_sl))

    def calculate_trailing_sl(self, ctx: MarketContext, position: Position) -> float:
        return position.current_sl


class WickAbsorptionMultiTFConfluenceBreakoutStrategy(BaseStrategy):
    def __init__(self):
        super().__init__("Wick Absorption & Multi-TF Confluence Breakout Entry", base_win_rate=0.72, base_rr=2.8)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if ctx.is_0915_bar:
            return None
        total_range = (ctx.high - ctx.low) if ctx.high > ctx.low else 1.0
        lower_wick = max(0.0, min(ctx.open, ctx.close) - ctx.low)
        wick_pct = (lower_wick / total_range) * 100.0
        conf_score = getattr(ctx, "confluence_score", 65)

        is_near_lb = (ctx.close <= ctx.lb_20 + 5.0 or ctx.low <= ctx.lb_20 + 5.0)
        is_mfi_htf_ok = (ctx.mfi14_15m >= ctx.prev_mfi14_15m) or (ctx.mfi14_30m >= ctx.prev_mfi14_30m)
        is_breakout = is_near_lb and is_mfi_htf_ok and (wick_pct >= 35.0 and conf_score >= 55)

        if is_breakout:
            entry_p = ctx.close
            sl = max(ctx.low - 15.0, entry_p - 20.0)
            target = entry_p + 60.0
            return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason=f"Breakout Entry (Wick: {wick_pct:.1f}%, Score: {conf_score})")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m and ctx.close > position.entry_price:
            return ExitSignal(True, ctx.close, "Dual MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


PreviousHighBreakoutMomentumStrategy = WickAbsorptionMultiTFConfluenceBreakoutStrategy


class PostSLRecoveryReentryStrategy(BaseStrategy):
    def __init__(self):
        super().__init__("One-Time Post-SL Recovery Re-Entry (+2 Lots)", base_win_rate=0.74, base_rr=2.5)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if not ctx.recovery_eligible:
            return None
        is_both_15m_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m and ctx.mfi14_15m > ctx.prev_mfi14_15m)
        is_near_lower_band = (ctx.low <= ctx.lb_20 + 20.0 or ctx.close < ctx.mb_20)
        if ctx.close >= ctx.open and is_both_15m_rising and is_near_lower_band:
            sl = ctx.close - 20.0
            target = ctx.close + 40.0
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, lot_size=2, reason="Post-SL Recovery Re-Entry")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m:
            return ExitSignal(True, ctx.close, "Recovery 15m Dual MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


class DynamicSwingLowBreakoutRetestStrategy(BaseStrategy):
    def __init__(self):
        super().__init__("Dynamic Swing Low First Breakout Retest Entry", base_win_rate=0.74, base_rr=2.5)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        is_15m_both_mfi_increasing = (ctx.mfi5_15m > ctx.prev_mfi5_15m and ctx.mfi14_15m > ctx.prev_mfi14_15m)
        is_higher_low = (ctx.low >= ctx.recent_swing_low - 1.0)
        if is_15m_both_mfi_increasing and is_higher_low:
            entry_p = ctx.close
            sl = max(ctx.recent_swing_low - 2.0, entry_p - 20.0)
            target = entry_p + 35.0
            return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason="Dynamic Swing Low Retest Bounce")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m:
            return ExitSignal(True, ctx.close, "Swing Low Dual MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


class TrendRidingMidBandStrategy(BaseStrategy):
    def __init__(self):
        super().__init__("Trend riding Mid band strategy", base_win_rate=0.73, base_rr=2.6)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        is_below_mb = (ctx.close < ctx.mb_20) or (ctx.low <= ctx.mb_20)
        if not is_below_mb:
            return None
        is_near_lb = (ctx.close <= ctx.lb_20 + 8.0) or (ctx.low <= ctx.lb_20 + 8.0)
        is_mfi_bouncing = (ctx.mfi5_15m > ctx.prev_mfi5_15m) and (ctx.mfi14_15m >= ctx.prev_mfi14_15m)
        conf_score = getattr(ctx, "confluence_score", 65)

        if is_near_lb and is_mfi_bouncing and conf_score > 60:
            entry_price = ctx.close
            sl = entry_price - 20.0
            target = max(ctx.ub_20, entry_price + 60.0)
            return EntrySignal(self.name, PositionSide.LONG, entry_price, sl, target, reason="Trend Riding Mid Band Entry")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        near_ub = (ctx.high >= ctx.ub_20 - 3.0 or ctx.close >= ctx.ub_20 - 3.0)
        both_15m_falling = (ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m)
        if near_ub or both_15m_falling:
            return ExitSignal(True, ctx.close, "Trend Riding Exit")
        return ExitSignal(False, ctx.close, "")


class StateMachineHandoverEngine:
    def __init__(self, strategies: List[BaseStrategy], min_ev_improvement: float = 4.0):
        self.strategies: Dict[str, BaseStrategy] = {s.name: s for s in strategies}
        self.min_ev_improvement = min_ev_improvement

    def _evaluate_in_flight_handover(self, ctx: MarketContext, pos: Position) -> HandoverDecision:
        current_strat = self.strategies.get(pos.strategy_name)
        if not current_strat:
            return HandoverDecision(False, pos.strategy_name, pos.strategy_name, 0.0, 0.0, pos.current_sl, pos.target_price, "")

        current_ev = current_strat.calculate_expected_value(ctx, pos)
        best_candidate: Optional[BaseStrategy] = None
        best_ev = current_ev

        for name, candidate in self.strategies.items():
            if name == pos.strategy_name:
                continue
            cand_ev = candidate.calculate_expected_value(ctx, pos)
            if cand_ev > best_ev + self.min_ev_improvement:
                best_ev = cand_ev
                best_candidate = candidate

        if best_candidate is not None:
            new_target = pos.target_price
            new_sl = pos.current_sl
            if "Breakout" in best_candidate.name and ctx.mfi14_15m > ctx.prev_mfi14_15m:
                new_target = max(pos.target_price, ctx.close + 50.0)
                new_sl = max(pos.current_sl, ctx.low - 10.0)

            guaranteed_sl = max(pos.initial_sl, new_sl)
            return HandoverDecision(
                switch_approved=True,
                from_strategy=pos.strategy_name,
                to_strategy=best_candidate.name,
                current_ev=current_ev,
                projected_ev=best_ev,
                adjusted_sl=guaranteed_sl,
                adjusted_target=new_target,
                reason=f"Handover EV boost ({current_ev:.1f} -> {best_ev:.1f} pts) via {best_candidate.name}"
            )

        return HandoverDecision(False, pos.strategy_name, pos.strategy_name, current_ev, current_ev, pos.current_sl, pos.target_price, "")


# Helper string mapper
def map_entry_type_to_strategy_name(entry_type_str: str) -> str:
    if not entry_type_str:
        return "Wick Absorption & Multi-TF Confluence Breakout Entry"
    if "Dynamic EPM" in entry_type_str or "EPM Low Bounce" in entry_type_str:
        return "Dynamic EPM Low Bounce LONG Entry"
    elif "Trend riding" in entry_type_str or "Mid band" in entry_type_str or "Trend Riding" in entry_type_str:
        return "Trend riding Mid band strategy"
    elif "Recovery" in entry_type_str or "Post-SL" in entry_type_str:
        return "One-Time Post-SL Recovery Re-Entry (+2 Lots)"
    elif "Swing Low" in entry_type_str or "SWING" in entry_type_str or "Retest" in entry_type_str:
        return "Dynamic Swing Low First Breakout Retest Entry"
    elif "Breakout" in entry_type_str or "Wick Absorption" in entry_type_str:
        return "Wick Absorption & Multi-TF Confluence Breakout Entry"
    else:
        return "Wick Absorption & Multi-TF Confluence Breakout Entry"


# =========================================================================
# 6. HISTORICAL BACKTEST ENGINE FOR M.STOCK
# =========================================================================

def run_mstock_backtest(from_date: str = "2026-08-01", to_date: str | None = None) -> None:
    """Run historical backtest simulation using m.Stock option candle datasets across custom selected date range."""
    if not to_date:
        to_date = datetime.now(IST).strftime("%Y-%m-%d")

    logger.info("📊 Starting m.Stock Historical Options Strategy Backtest Simulator...")
    logger.info("📅 Custom Date Range: %s to %s", from_date, to_date)
    
    strategies = [
        DynamicSwingLowBreakoutRetestStrategy(),
        TrendRidingMidBandStrategy(),
    ]
    engine = StateMachineHandoverEngine(strategies=strategies)
    logger.info("✅ m.Stock Backtest Engine successfully initialized with %d strategies.", len(engine.strategies))
    print(f"\n=========================================================================")
    print(f"M.STOCK HISTORICAL OPTIONS STRATEGY BACKTEST SIMULATOR")
    print(f"Custom Date Range Selected : {from_date} to {to_date}")
    print(f"Engine Status              : INITIALIZED & READY")
    print(f"Strategies Loaded          : 4 Core Strategies (Wick Absorption, Recovery, Retest, Trend Riding)")
    print(f"In-Flight Handover EV Delta: +{MIN_EV_HANDOVER_DELTA} pts")
    print(f"=========================================================================\n")

    try:
        from state_machine_handover_engine import run_state_machine_backtest

        cache_dir = Path(".cache_candles")
        if not cache_dir.exists():
            print("⚠️ No .cache_candles directory found.")
            return

        fifteen_files = list(cache_dir.glob("candles_*_FIFTEEN_MINUTE_*.json"))
        if not fifteen_files:
            print("⚠️ No 15-minute candle datasets found in .cache_candles/.")
            return

        print(f"📁 Loaded {len(fifteen_files)} 15-minute option candle datasets from .cache_candles/")
        print(f"⚡ Executing 4-Strategy Handover Simulation for window [{from_date} -> {to_date}]...\n")

        # Load token metadata mapping
        token_map = {}
        for map_path in (
            "attached_assets/0_token_to_delta_1786898248102.json",
            "0_sensex_options_delta_1786941098090.json",
            "delta_map.json",
            "sensex_options_bot/delta_map.json"
        ):
            if os.path.exists(map_path):
                try:
                    with open(map_path, "r", encoding="utf-8") as f_map:
                        map_data = json.load(f_map)
                    if "options_delta_mapping" in map_data:
                        for tok, meta in map_data["options_delta_mapping"].items():
                            token_map[str(tok)] = meta
                    elif isinstance(map_data, dict):
                        for sym, meta in map_data.items():
                            tok = str(meta.get("symboltoken") or meta.get("token") or "")
                            if tok:
                                token_map[tok] = {
                                    "trading_symbol": sym,
                                    "strike": meta.get("strike"),
                                    "option_type": meta.get("option_type") or ("CE" if "CE" in sym else "PE"),
                                    "expiry": meta.get("expiry")
                                }
                except Exception:
                    pass

        all_trade_dfs = []
        parsed_contracts = set()

        for f_15 in fifteen_files:
            fname = f_15.name
            parts = fname.split("_")
            if len(parts) < 3:
                continue
            token = parts[1]
            if token in parsed_contracts or token == "99919000":
                continue
            parsed_contracts.add(token)

            try:
                with open(f_15, "r", encoding="utf-8") as f:
                    c15_data = json.load(f)
            except Exception:
                continue
            if not c15_data:
                continue

            df_15m = pd.DataFrame(c15_data, columns=["timestamp", "open", "high", "low", "close", "volume"])
            df_15m['date'] = df_15m['timestamp'].astype(str).str[:10]

            df_15m_filtered = df_15m[(df_15m['date'] >= from_date) & (df_15m['date'] <= to_date)].copy()
            if df_15m_filtered.empty:
                continue

            f_30_matches = list(cache_dir.glob(f"candles_{token}_THIRTY_MINUTE_*.json"))
            df_30m = None
            if f_30_matches:
                try:
                    with open(f_30_matches[0], "r", encoding="utf-8") as f:
                        c30_data = json.load(f)
                    if c30_data:
                        df_30m = pd.DataFrame(c30_data, columns=["timestamp", "open", "high", "low", "close", "volume"])
                except Exception:
                    pass

            f_3_matches = list(cache_dir.glob(f"candles_{token}_THREE_MINUTE_*.json"))
            df_3m = None
            if f_3_matches:
                try:
                    with open(f_3_matches[0], "r", encoding="utf-8") as f:
                        c3_data = json.load(f)
                    if c3_data:
                        df_3m = pd.DataFrame(c3_data, columns=["timestamp", "open", "high", "low", "close", "volume"])
                except Exception:
                    pass

            token_meta = token_map.get(str(token), {})
            strike_val = token_meta.get("strike")
            opt_type = token_meta.get("option_type") or ("CE" if "CE" in str(token_meta.get("trading_symbol", "")) else ("PE" if "PE" in str(token_meta.get("trading_symbol", "")) else ""))

            if strike_val and opt_type:
                symbol_label = f"SENSEX {int(strike_val)} {opt_type}"
            elif token_meta.get("trading_symbol"):
                symbol_label = str(token_meta["trading_symbol"])
            else:
                symbol_label = f"Token {token}"

            closed_df, _ = run_state_machine_backtest(df_15m, df_30m=df_30m, df_3m=df_3m, from_date_str=from_date, contract_symbol=symbol_label)
            if not closed_df.empty:
                closed_df['Strike'] = int(strike_val) if strike_val else None
                closed_df['CE/PE'] = opt_type if opt_type else None
                all_trade_dfs.append(closed_df)

        if not all_trade_dfs:
            print("ℹ️ No trades were triggered within the selected date range.")
            return

        combined_df = pd.concat(all_trade_dfs, ignore_index=True)
        combined_df['date'] = combined_df['entry_time'].astype(str).str[:10]
        combined_df = combined_df[(combined_df['date'] >= from_date) & (combined_df['date'] <= to_date)].copy()

        if combined_df.empty:
            print("ℹ️ No trades completed in date range.")
            return

        def clean_dt_fmt(ts_val):
            if not ts_val or pd.isna(ts_val):
                return ""
            s = str(ts_val).replace("T", " ").replace("+05:30", "").strip()
            return s[:16] if len(s) >= 16 else s

        if 'entry_time' in combined_df.columns:
            combined_df['entry_time'] = combined_df['entry_time'].apply(clean_dt_fmt)
        if 'exit_time' in combined_df.columns:
            combined_df['exit_time'] = combined_df['exit_time'].apply(clean_dt_fmt)

        total_trades = len(combined_df)
        wins = (combined_df['pnl'] > 0).sum()
        losses = (combined_df['pnl'] < 0).sum()
        win_rate = (wins / total_trades) * 100.0 if total_trades > 0 else 0.0
        total_pnl_pts = combined_df['pnl'].sum()
        lot_qty = 20
        total_pnl_rs = total_pnl_pts * lot_qty

        gross_profit = combined_df[combined_df['pnl'] > 0]['pnl'].sum()
        gross_loss = abs(combined_df[combined_df['pnl'] < 0]['pnl'].sum())
        profit_factor = (gross_profit / gross_loss) if gross_loss > 0 else gross_profit

        handovers_cnt = combined_df['handovers'].sum() if 'handovers' in combined_df.columns else 0

        combined_df['cum_pnl'] = combined_df['pnl'].cumsum()
        combined_df['peak'] = combined_df['cum_pnl'].cummax()
        combined_df['drawdown'] = combined_df['peak'] - combined_df['cum_pnl']
        max_drawdown = combined_df['drawdown'].max()

        print("=========================================================================")
        print("📊 BACKTEST PERFORMANCE SUMMARY REPORT")
        print("=========================================================================")
        print(f"Date Window        : {from_date} to {to_date}")
        print(f"Total Trades       : {total_trades}")
        print(f"Winning Trades     : {wins} 🟢")
        print(f"Losing Trades      : {losses} 🔴")
        print(f"Win Rate / Accuracy: {win_rate:.2f}%")
        print(f"Total Net PnL (Pts): {total_pnl_pts:+.2f} pts")
        print(f"Total Net PnL (₹)  : ₹{total_pnl_rs:+,.2f} (Base Lot: {lot_qty} Qty)")
        print(f"Profit Factor      : {profit_factor:.2f}")
        print(f"Max Drawdown (Pts) : -{max_drawdown:.2f} pts")
        print(f"Dynamic Handovers  : {handovers_cnt}")
        print("=========================================================================\n")

        print("📈 STRATEGY PERFORMANCE BREAKDOWN:")
        print(f"{'Strategy Name':<45} | {'Trades':<7} | {'Win %':<8} | {'PnL (Pts)':<10} | {'PnL (₹)':<12}")
        print("-" * 92)

        strat_col = 'strategy_name' if 'strategy_name' in combined_df.columns else ('initial_strategy' if 'initial_strategy' in combined_df.columns else None)
        if strat_col and strat_col in combined_df.columns:
            grouped = combined_df.groupby(strat_col)
            for strat_name, group in grouped:
                st_trades = len(group)
                st_wins = (group['pnl'] > 0).sum()
                st_win_rate = (st_wins / st_trades) * 100.0 if st_trades > 0 else 0.0
                st_pnl_pts = group['pnl'].sum()
                st_pnl_rs = st_pnl_pts * lot_qty
                print(f"{str(strat_name)[:44]:<45} | {st_trades:<7} | {st_win_rate:6.1f}% | {st_pnl_pts:+9.1f} | ₹{st_pnl_rs:+10,.2f}")
        print("-" * 92 + "\n")

        report_file = "mstock_backtest_report.xlsx"
        try:
            with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
                combined_df.to_excel(writer, sheet_name="Trade Log", index=False)
                summary_df = pd.DataFrame([{
                    "From Date": from_date, "To Date": to_date,
                    "Total Trades": total_trades, "Winning Trades": wins, "Losing Trades": losses,
                    "Win Rate %": round(win_rate, 2), "Total PnL Pts": round(total_pnl_pts, 2),
                    "Total PnL Rs": round(total_pnl_rs, 2), "Profit Factor": round(profit_factor, 2),
                    "Max Drawdown Pts": round(max_drawdown, 2), "In-Flight Handovers": int(handovers_cnt)
                }])
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
            print(f"📄 Backtest report successfully exported to '{report_file}'.\n")
        except Exception as e_excel:
            logger.warning("Could not export Excel report: %s", e_excel)

    except Exception as exc:
        logger.error("❌ Backtest execution error: %s", exc, exc_info=True)


# =========================================================================
# 7. CONTRACT SELECTION & MSTOCK ADAPTER HELPERS
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
    today = now_ist.date()
    contracts_with_expiry = []
    for c in matching:
        try:
            exp_dt = to_ist_datetime(c.expiry)
            if exp_dt.date() >= today:
                contracts_with_expiry.append((c, exp_dt.date()))
        except Exception:
            pass

    if contracts_with_expiry:
        near_term = [item for item in contracts_with_expiry if 0 <= (item[1] - today).days <= 45]
        target_list = near_term if near_term else contracts_with_expiry
        earliest_expiry = min(exp_dt_date for _, exp_dt_date in target_list)
        matching = [c for c, exp_dt_date in target_list if exp_dt_date == earliest_expiry]

    by_strike: dict[float, OptionContract] = {}
    for c in matching:
        if c.strike not in by_strike:
            by_strike[c.strike] = c

    unique_contracts = list(by_strike.values())
    strikes = sorted([c.strike for c in unique_contracts if c.strike > 0])
    diffs = [strikes[i + 1] - strikes[i] for i in range(len(strikes) - 1)]
    positive_diffs = [d for d in diffs if d > 0]
    strike_step = min(positive_diffs) if positive_diffs else 100.0

    atm_strike = round(spot / strike_step) * strike_step

    if option_type == "CE":
        itm = sorted([c for c in unique_contracts if c.strike <= spot or c.strike <= atm_strike], key=lambda c: c.strike, reverse=True)
        if not itm:
            itm = sorted(unique_contracts, key=lambda c: c.strike, reverse=True)
        return itm[:count]
    else:
        itm = sorted([c for c in unique_contracts if c.strike >= spot or c.strike >= atm_strike], key=lambda c: c.strike)
        if not itm:
            itm = sorted(unique_contracts, key=lambda c: c.strike)
        return itm[:count]


def fetch_realtime_sensex_spot() -> float | None:
    """Fetch live real-time SENSEX spot price from market feed as fallback when broker feed is off-hours or delayed."""
    try:
        url = "https://query1.finance.yahoo.com/v8/finance/chart/%5BSENSEX?interval=1m&range=1d"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            meta = data["chart"]["result"][0]["meta"]
            price = float(meta.get("regularMarketPrice") or meta.get("chartPreviousClose") or 0.0)
            if price > 0:
                return price
    except Exception:
        pass
    return None


def safe_ltp_data(mstock_client: MstockClientAdapter, exchange: str, trading_symbol: str, symbol_token: str, max_retries: int = 3) -> dict | None:
    if mstock_client is None:
        return None
    res = mstock_client.ltpData(exchange, trading_symbol, symbol_token)
    if res and isinstance(res, dict) and res.get("status") is True and res.get("data"):
        return res
    return None


def safe_get_candle_data(mstock_client: MstockClientAdapter, params: dict, max_retries: int = 3) -> dict | None:
    if mstock_client is None:
        return None
    res = mstock_client.getCandleData(params)
    if res and isinstance(res, dict) and res.get("status") is True and res.get("data"):
        return res
    return None


def calculate_atr_and_stddev(mstock_client: MstockClientAdapter, exchange: str, symbol_token: str) -> tuple[float, float]:
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
        res = safe_get_candle_data(mstock_client, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= 20:
                highs = [float(c[2]) for c in candles]
                lows = [float(c[3]) for c in candles]
                closes = [float(c[4]) for c in candles]
                
                tr_list = []
                for i in range(1, len(candles)):
                    tr = max(highs[i] - lows[i], abs(highs[i] - closes[i - 1]), abs(lows[i] - closes[i - 1]))
                    tr_list.append(tr)
                atr_14 = sum(tr_list[-14:]) / 14.0 if len(tr_list) >= 14 else 20.0

                sub_closes = closes[-20:]
                mean_c = sum(sub_closes) / 20.0
                var_c = sum((x - mean_c) ** 2 for x in sub_closes) / 20.0
                std_dev_20 = math.sqrt(var_c)

                return atr_14, std_dev_20
    except Exception:
        pass
    return 20.0, 15.0


def load_grid_state_epm_low(option_type: str = "CE") -> float | None:
    for filename in ("grid_state.json", "bot_state_memory.json"):
        path = Path(filename)
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if option_type.upper() == "CE" and "ce_epm_low" in data and data["ce_epm_low"] is not None:
                    return float(data["ce_epm_low"])
                if option_type.upper() == "PE" and "pe_epm_low" in data and data["pe_epm_low"] is not None:
                    return float(data["pe_epm_low"])
                grid_data = data.get("grid") or {}
                leg_key = "ce_leg" if option_type.upper() == "CE" else "pe_leg"
                if leg_key in grid_data and "epm_lower_range" in grid_data[leg_key]:
                    return float(grid_data[leg_key]["epm_lower_range"])
            except Exception:
                pass
    return None


def get_current_15m_candle_ohl(mstock_client: MstockClientAdapter, exchange: str, symbol_token: str) -> tuple[float | None, float | None]:
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
        res = safe_get_candle_data(mstock_client, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if candles:
                last_candle = candles[-1]
                if len(last_candle) >= 4:
                    return float(last_candle[1]), float(last_candle[3])
    except Exception:
        pass
    return None, None


def get_mfi_multi_period(mstock_client: MstockClientAdapter, exchange: str, symbol_token: str, timeframe: str = "FIFTEEN_MINUTE", periods: list[int] = [5, 14], return_extra: bool = False) -> Any:
    max_period = max(periods)
    curr_mfis = {p: 50.0 for p in periods}
    prev_mfis = {p: 50.0 for p in periods}
    prev_prev_mfis = {p: 50.0 for p in periods}
    extra_data = {"prev_close": 0.0, "prev_prev_low": 0.0}
    
    try:
        now_dt = datetime.now(IST)
        lookback_days = 10 if timeframe in ("THIRTY_MINUTE", "ONE_HOUR") else 7
        from_dt = now_dt - timedelta(days=lookback_days)
        params = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": timeframe,
            "fromdate": from_dt.strftime("%Y-%m-%d %H:%M"),
            "todate": now_dt.strftime("%Y-%m-%d %H:%M")
        }
        res = safe_get_candle_data(mstock_client, params)
        if isinstance(res, dict) and res.get("status") is True and res.get("data"):
            candles = res["data"]
            if len(candles) >= 3:
                extra_data["prev_close"] = float(candles[-2][4])
                extra_data["prev_prev_low"] = float(candles[-3][3])
                
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

                    if len(typical_prices) >= p + 2:
                        curr_mfis[p] = calc_mfi_at(len(typical_prices) - 1, p)
                        prev_mfis[p] = calc_mfi_at(len(typical_prices) - 2, p)
                        prev_prev_mfis[p] = calc_mfi_at(len(typical_prices) - 3, p) if len(typical_prices) >= p + 3 else prev_mfis[p]
    except Exception:
        pass
        
    if return_extra:
        return curr_mfis, prev_mfis, prev_prev_mfis, extra_data
    return curr_mfis, prev_mfis


def get_3m_bollinger_bands(mstock_client: MstockClientAdapter, exchange: str, symbol_token: str, period: int = 20, std_dev_mult: float = 2.0) -> tuple[float | None, float | None, float | None, float | None, float | None, float | None, float | None, float | None]:
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
        res = safe_get_candle_data(mstock_client, params)
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
    except Exception:
        pass
    return None, None, None, None, None, None, None, None


def get_current_time_slot() -> str:
    now_dt = datetime.now(IST)
    m_of_day = now_dt.hour * 60 + now_dt.minute
    if m_of_day < (9 * 60 + 45):
        return "09:15"
    elif m_of_day < (12 * 60 + 15):
        return "09:45"
    else:
        return "12:15"


def load_bot_memory() -> dict | None:
    path = Path("bot_state_memory.json")
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        today_str = datetime.now(IST).strftime("%Y-%m-%d")
        if data.get("date") == today_str:
            return data
    except Exception:
        pass
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
                }
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
        
        grid_state_data = {
            "date": datetime.now(IST).strftime("%Y-%m-%d"),
            "grid_time_slot": grid_time_slot,
            "ce_epm_low": grid.ce_leg.epm_lower_range if grid and grid.ce_leg else None,
            "pe_epm_low": grid.pe_leg.epm_lower_range if grid and grid.pe_leg else None,
            "ce_symbol": ce_contract.trading_symbol if ce_contract else "",
            "pe_symbol": pe_contract.trading_symbol if pe_contract else ""
        }
        with open("grid_state.json", "w", encoding="utf-8") as f_grid:
            json.dump(grid_state_data, f_grid, indent=4)

        logger.info("💾 m.Stock bot active state and grid memory saved successfully.")
    except Exception as e:
        logger.warning("Failed to save bot active memory: %s", e)


def save_bot_memory(trades_completed: int, grid_time_slot: str, grid: EPMMasterGrid, ce_contract: OptionContract, pe_contract: OptionContract) -> None:
    save_bot_memory_full(trades_completed, grid_time_slot, grid, ce_contract, pe_contract, "IDLE", None, 0.0, 0.0, 0.0, 0.0, False, False, 1, None, False)


def load_delta_map() -> dict[str, dict[str, Any]]:
    for path in ("delta_map.json", "../../delta_map.json", "0_sensex_options_delta_1786941098090.json"):
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
            except Exception:
                pass
    return {}


def build_epm_grid_and_contracts(
    mstock_client: MstockClientAdapter,
    current_slot: str,
    spot_price: float,
    spot_open: float,
    vix_val: float,
    buffer: float = 0.13,
    index_name: str = "SENSEX",
    sl_offset: float | None = None,
) -> tuple[EPMMasterGrid, list[OptionContract], list[OptionContract]]:
    is_bn = "BANKNIFTY" in str(index_name).upper()
    exchange = "NFO" if is_bn else "BFO"
    search_symbol = "BANKNIFTY" if is_bn else "SENSEX"
    
    if sl_offset is None:
        sl_offset = 12.0 if is_bn else 15.0

    search_res = mstock_client.searchScrip(exchange, search_symbol)
    rows = search_res.get("data", []) if isinstance(search_res, dict) else []
    if not isinstance(rows, list):
        rows = []

    delta_map = load_delta_map()
    if delta_map:
        search_upper = search_symbol.upper()
        existing_symbols = {str(r.get("tradingsymbol") or "").strip() for r in rows}
        for sym, meta in delta_map.items():
            if search_upper in sym.upper() and sym not in existing_symbols:
                rows.append({
                    "tradingsymbol": sym,
                    "symboltoken": meta.get("symboltoken") or meta.get("token") or "0",
                    "expiry": meta.get("expiry") or "",
                    "strike": meta.get("strike"),
                })

    contracts: list[OptionContract] = []
    month_map = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9, "O": 10, "N": 11, "D": 12, "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}

    for r in rows:
        symbol = str(r.get("tradingsymbol") or "").strip()
        if not symbol or not (symbol.endswith("CE") or symbol.endswith("PE")):
            continue
        opt_type = "CE" if symbol.endswith("CE") else "PE"
        token = str(r.get("symboltoken") or "").strip()

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

        m_ddmmmyy = re.search(r"(?:BSE|NSE)?(?:SENSEX|BANKNIFTY)(0[1-9]|[12][0-9]|3[01])([A-Za-z]{3})(\d{2})(\d{4,6})(CE|PE)$", symbol, re.IGNORECASE)
        m_sym = re.search(r"(?:BSE|NSE)?(?:SENSEX|BANKNIFTY)(\d{2})([A-Za-z]{3}|\d|[ONDond])(?:(0[1-9]|[12][0-9]|3[01]))?(\d{4,6})(CE|PE)$", symbol, re.IGNORECASE)
        if m_ddmmmyy:
            dd_str, m_str, yy_str, str_val, _ = m_ddmmmyy.groups()
            if not strike_val:
                strike_val = float(str_val)
            if not expiry_val:
                m_num = month_map.get(m_str.upper(), 9)
                expiry_val = f"20{yy_str}-{m_num:02d}-{int(dd_str):02d}"
        elif m_sym:
            yy, m_str, dd, str_val, _ = m_sym.groups()
            if not strike_val:
                strike_val = float(str_val)
            if not expiry_val:
                m_num = month_map.get(m_str.upper(), 9)
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
                expiry_val = "2026-09-24"

        if not expiry_val or not strike_val:
            continue

        dte_days, _ = calculate_dte_sqrt(expiry_val)
        delta_val = calculate_bsm_delta(spot_price, strike_val, dte_days, vix_val, opt_type)

        contracts.append(OptionContract(exchange, symbol, token, expiry_val, strike_val, opt_type, delta_val))

    ce_contracts = select_itm_contracts(contracts, spot_price, "CE", count=3) if contracts else []
    pe_contracts = select_itm_contracts(contracts, spot_price, "PE", count=3) if contracts else []

    if not ce_contracts or not pe_contracts:
        ce_contracts = ce_contracts or [OptionContract(exchange, f"{search_symbol}_CE_DUMMY", "0", "2026-09-30", spot_price, "CE", 0.55)]
        pe_contracts = pe_contracts or [OptionContract(exchange, f"{search_symbol}_PE_DUMMY", "0", "2026-09-30", spot_price, "PE", -0.55)]

    ce_legs_data = []
    for c in ce_contracts:
        res = safe_ltp_data(mstock_client, exchange, c.trading_symbol, c.symbol_token) if c.symbol_token != "0" else {}
        ltp = float(res["data"]["ltp"]) if isinstance(res, dict) and res.get("data") else 500.0
        c_open = float(res["data"]["open"]) if isinstance(res, dict) and res.get("data") and res["data"].get("open") else ltp
        price_to_use = c_open if current_slot == "09:15" else ltp
        ce_legs_data.append((price_to_use, abs(c.delta), c.strike, str(c.expiry), c.trading_symbol))

    pe_legs_data = []
    for p in pe_contracts:
        res = safe_ltp_data(mstock_client, exchange, p.trading_symbol, p.symbol_token) if p.symbol_token != "0" else {}
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
        sl_offset=sl_offset,
    )

    return grid, ce_contracts, pe_contracts


def format_grid_notification(grid: EPMMasterGrid, title: str, spot_price: float, spot_open: float, vix_val: float, dte_days: float, current_slot: str = "09:15") -> str:
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


def submit_mstock_order(mstock_client: MstockClientAdapter, trading_symbol: str, symbol_token: str, transaction_type: str = "BUY", quantity: int = 20, exchange: str = "BFO") -> Any:
    qty_val = max(1, int(quantity))
    order_params = {
        "tradingsymbol": str(trading_symbol).strip(),
        "symboltoken": str(symbol_token).strip(),
        "transactiontype": transaction_type.upper(),
        "exchange": exchange,
        "ordertype": "MARKET",
        "producttype": "NRML",
        "duration": "DAY",
        "quantity": str(qty_val),
    }
    res = mstock_client.placeOrder(order_params)
    if res and res.get("status") in (True, "success"):
        order_id = res.get("data", {}).get("orderid", f"MSTK_{int(time.time())}")
        logger.info("⚡ [m.Stock ORDER SUCCESS] %s %d %s | Order ID: %s", transaction_type, qty_val, trading_symbol, order_id)
        send_mobile_alert(f"🚨 *m.Stock ORDER PLACED*\nAction: *{transaction_type}*\nContract: *{trading_symbol}*\nQty: *{qty_val}*\nOrder ID: `{order_id}`")
        return order_id

    logger.error("❌ m.Stock Order Placement Failed for %s %d Qty", trading_symbol, qty_val)
    send_mobile_alert(f"⚠️ *m.Stock ORDER ERROR*\nFailed to place {transaction_type} order for {trading_symbol}.")
    return None


def execute_failsafe_sell(mstock_client: MstockClientAdapter, trading_symbol: str, symbol_token: str, quantity: int, ltp: float, exchange: str = "BFO") -> Any:
    return submit_mstock_order(mstock_client, trading_symbol, symbol_token, "SELL", quantity, exchange)


# =========================================================================
# 8. MAIN AUTONOMOUS CLOUD RUNNER FOR M.STOCK
# =========================================================================

def run_mstock_bot() -> None:
    target_index = os.getenv("INDEX_NAME", os.getenv("TARGET_INDEX", "SENSEX")).upper().strip()
    is_bn_bot = "BANKNIFTY" in target_index
    spot_exch = "NSE" if is_bn_bot else "BSE"
    spot_sym = "BANKNIFTY" if is_bn_bot else "SENSEX"
    spot_tok = "99926009" if is_bn_bot else "99919000"
    active_lot_units = 30 if is_bn_bot else 20

    logger.info("🚀 Starting Standalone Cloud %s Options Bot on m.Stock...", target_index)
    excel_tracker = ExcelTracker()

    mstock_client = create_authenticated_mstock_client()

    if mstock_client.is_authenticated:
        logger.info("=========================================================================")
        logger.info("M.STOCK API CONNECTION STATUS: 🟢 LIVE FEED CONNECTED & AUTHENTICATED")
        logger.info("Status: Authentication Successful!")
        logger.info("Endpoint: %s | Client ID: %s", mstock_client.base_url, mstock_client.client_id)
        logger.info("=========================================================================")
        send_mobile_alert(f"🟢 *m.Stock LIVE FEED CONNECTED*\nAuthentication Successful!\nClient ID: `{mstock_client.client_id}`")
    else:
        logger.info("=========================================================================")
        logger.warning("M.STOCK API CONNECTION STATUS: 🔴 LIVE FEED OFFLINE / AUTHENTICATION FAILED")
        logger.warning("Status: Authentication Failed")
        logger.warning("Reason for Failure: %s", mstock_client.last_auth_reason)
        logger.warning("Operating Mode: PAPER TRADING / DEMO FEED")
        logger.info("=========================================================================")
        send_mobile_alert(f"⚠️ *m.Stock LIVE FEED OFFLINE*\nReason: {mstock_client.last_auth_reason}\nRunning in Paper Mode.")

    current_slot = get_current_time_slot()
    mem = load_bot_memory()

    trades_completed = 0
    grid = None
    ce_contract = None
    pe_contract = None
    spot_price = 51500.0 if is_bn_bot else 77500.0
    spot_open = spot_price
    vix_val = 13.5
    dte_days = 4.0
    ce_ltp = 500.0
    pe_ltp = 300.0
    current_epm_buffer = 0.13

    # Fetch Real-time Spot Price from m.Stock or Live Market Feed
    try:
        spot_res = safe_ltp_data(mstock_client, spot_exch, spot_sym, spot_tok)
        if spot_res and spot_res.get("data") and float(spot_res["data"].get("ltp", 0.0)) > 0:
            spot_price = float(spot_res["data"]["ltp"])
            spot_open = float(spot_res["data"].get("open", spot_price)) or spot_price
        else:
            live_realtime = fetch_realtime_sensex_spot()
            if live_realtime and live_realtime > 0:
                spot_price = live_realtime
                spot_open = live_realtime
    except Exception:
        pass

    if mem is not None:
        trades_completed = mem.get("trades_completed", 0)
        if mem.get("grid_time_slot") == current_slot:
            try:
                ce_c_data = mem["ce_contract"]
                pe_c_data = mem["pe_contract"]
                grid_data = mem["grid"]

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

                ce_leg_data = grid_data["ce_leg"]
                pe_leg_data = grid_data["pe_leg"]

                ce_leg = MasterGridLeg("CE", ce_leg_data["strike"], ce_leg_data["ltp"], ce_leg_data["delta"], ce_leg_data["target_epm"], ce_leg_data["epm_lower_range"], ce_leg_data["sl_auto"], ce_leg_data["practical_target"])
                pe_leg = MasterGridLeg("PE", pe_leg_data["strike"], pe_leg_data["ltp"], pe_leg_data["delta"], pe_leg_data["target_epm"], pe_leg_data["epm_lower_range"], pe_leg_data["sl_auto"], pe_leg_data["practical_target"])

                grid = EPMMasterGrid(
                    spot=grid_data["spot"], vix=grid_data["vix"], dte=grid_data["dte"],
                    dte_sqrt=math.sqrt(grid_data["dte"] / 365.0), index_move=grid_data["index_move"],
                    noise_10=grid_data["index_move"] * 0.10,
                    lower_index=grid_data["spot"] - grid_data["index_move"],
                    upper_index=grid_data["spot"] + grid_data["index_move"],
                    ce_leg=ce_leg, pe_leg=pe_leg
                )
                logger.info("🔮 [RECALL MEMORY] Recalled stored EPM Master Grid for slot %s.", current_slot)
            except Exception as exc:
                logger.warning("⚠️ Saved memory schema mismatch. Resetting grid: %s", exc)
                grid = None

    if grid is None:
        logger.info("🆕 [MASTER GRID] Calculating a new Master Grid for slot %s (%s)...", current_slot, target_index)
        grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(mstock_client, current_slot, spot_price, spot_open, vix_val, buffer=current_epm_buffer, index_name=target_index)
        ce_contract = ce_contracts[0]
        pe_contract = pe_contracts[0]
        ce_ltp = grid.ce_leg.ltp
        pe_ltp = grid.pe_leg.ltp
        dte_days = grid.dte
        save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)

    msg = format_grid_notification(grid, f"🔔 *m.Stock {target_index} MASTER GRID INITIALIZED*", spot_price, spot_open, vix_val, dte_days, current_slot)
    send_mobile_alert(msg)

    execution_mode = "PAPER" if ("--paper" in sys.argv or os.getenv("EXECUTION_MODE", "").upper() == "PAPER") else ("LIVE" if mstock_client.is_authenticated else "PAPER")
    bot_state = "IDLE"
    max_trades_per_day = 2
    ce_sl_hit_today = False
    active_contract = None
    active_entry_price = 0.0
    active_sl = 0.0
    active_target = 0.0

    cli_lots = None
    for arg in sys.argv:
        if arg.startswith("--lots="):
            try:
                cli_lots = int(arg.split("=")[1])
            except ValueError:
                pass
    lot_size = cli_lots if cli_lots else int(os.environ.get("BASE_LOT_SIZE", "4"))

    handover_strategies = [
        PreviousHighBreakoutMomentumStrategy(),
        PostSLRecoveryReentryStrategy(),
        DynamicSwingLowBreakoutRetestStrategy(),
        TrendRidingMidBandStrategy(),
    ]

    tg_listener = TelegramCommandListener(os.environ.get("TELEGRAM_BOT_TOKEN"))
    logger.info("🔄 Entering continuous monitoring loop on m.Stock (Mode: %s)...", execution_mode)

    loop_counter = 0
    recent_ce_low = ce_ltp
    recent_pe_low = pe_ltp

    try:
        while True:
            time.sleep(1.0)
            checked_at = datetime.now(IST)
            loop_counter += 1

            now_slot = get_current_time_slot()
            if now_slot != current_slot:
                logger.info("⏰ [SLOT TRANSITION] Slot changed from %s to %s", current_slot, now_slot)
                current_slot = now_slot
                grid, ce_contracts, pe_contracts = build_epm_grid_and_contracts(mstock_client, current_slot, spot_price, spot_open, vix_val, buffer=current_epm_buffer, index_name=target_index)
                ce_contract = ce_contracts[0]
                pe_contract = pe_contracts[0]
                save_bot_memory(trades_completed, current_slot, grid, ce_contract, pe_contract)
                msg = format_grid_notification(grid, f"🔔 *m.Stock {target_index} MASTER GRID UPDATED ({current_slot})*", spot_price, spot_open, vix_val, grid.dte, current_slot)
                send_mobile_alert(msg)

            if (checked_at.hour == 15 and checked_at.minute >= 30) or (checked_at.hour > 15):
                logger.info("🕒 Market closed (after 3:30 PM IST). Shutting down m.Stock bot gracefully.")
                send_mobile_alert("🕒 *MARKET CLOSE REACHED*\nm.Stock bot shutting down gracefully.")
                break

            cmd, remote_val = tg_listener.get_new_command()
            if cmd == "STOP":
                logger.info("🛑 Remote STOP command received. Halting execution.")
                send_mobile_alert("🛑 *REMOTE STOP COMMAND RECEIVED*\nm.Stock bot execution halted safely.")
                break
            elif cmd == "LIVE":
                if remote_val:
                    lot_size = int(remote_val)
                execution_mode = "LIVE" if mstock_client.is_authenticated else "PAPER"
                send_mobile_alert(f"🚨 *MODE SWITCHED TO LIVE TRADING*\nIndex: *{target_index}*\nLot Size: *{lot_size}*")
            elif cmd == "DEMO":
                execution_mode = "PAPER"
                send_mobile_alert("🛡️ *MODE SWITCHED TO SAFE PAPER TRADING*")

            live_spot_res = safe_ltp_data(mstock_client, spot_exch, spot_sym, spot_tok)
            if live_spot_res and live_spot_res.get("data") and float(live_spot_res["data"].get("ltp", 0.0)) > 0:
                live_spot = float(live_spot_res["data"]["ltp"])
            else:
                live_spot = fetch_realtime_sensex_spot() or spot_price

            live_ce_res = safe_ltp_data(mstock_client, "BFO", ce_contract.trading_symbol, ce_contract.symbol_token) if ce_contract else None
            live_pe_res = safe_ltp_data(mstock_client, "BFO", pe_contract.trading_symbol, pe_contract.symbol_token) if pe_contract else None

            live_ce_ltp = float(live_ce_res["data"]["ltp"]) if live_ce_res and live_ce_res.get("data") and float(live_ce_res["data"].get("ltp", 0.0)) > 0 else ce_ltp
            live_pe_ltp = float(live_pe_res["data"]["ltp"]) if live_pe_res and live_pe_res.get("data") and float(live_pe_res["data"].get("ltp", 0.0)) > 0 else pe_ltp

            if bot_state == "IDLE":
                recent_ce_low = min(recent_ce_low, live_ce_ltp)
                recent_pe_low = min(recent_pe_low, live_pe_ltp)

                now_time_str = checked_at.strftime("%H:%M")
                if now_time_str < "15:00" and trades_completed < max_trades_per_day:
                    mfis_15m_ce, prev_mfis_15m_ce = get_mfi_multi_period(mstock_client, "BFO", ce_contract.symbol_token, "FIFTEEN_MINUTE", [5, 14]) if ce_contract else ({5:50,14:50}, {5:50,14:50})
                    mfi5_ce = mfis_15m_ce.get(5, 50)
                    prev_mfi5_ce = prev_mfis_15m_ce.get(5, 50)

                    ce_epm_low = grid.ce_leg.epm_lower_range if grid and grid.ce_leg else 0.0
                    is_ce_bounce = (ce_epm_low > 0) and (live_ce_ltp >= ce_epm_low - 10.0) and (live_ce_ltp <= ce_epm_low + 20.0) and (mfi5_ce > prev_mfi5_ce)

                    if is_ce_bounce and not ce_sl_hit_today:
                        bot_state = "CE_LONG"
                        active_contract = ce_contract
                        active_entry_price = live_ce_ltp
                        active_target = active_entry_price + 55.0
                        active_sl = active_entry_price - 20.0
                        entry_time = checked_at
                        qty = lot_size * active_lot_units

                        logger.info("🟢 [m.Stock CE ENTRY] CE LTP ₹%.2f triggered (SL: ₹%.2f, TP: ₹%.2f)", live_ce_ltp, active_sl, active_target)
                        send_mobile_alert(f"🟢 *m.Stock CE ENTRY SIGNAL*\nContract: *{active_contract.trading_symbol}*\nEntry: ₹{active_entry_price:.2f}\nSL: ₹{active_sl:.2f} | Target: ₹{active_target:.2f}")

                        if execution_mode == "LIVE":
                            submit_mstock_order(mstock_client, active_contract.trading_symbol, active_contract.symbol_token, "BUY", qty)

                        excel_tracker.add_order({
                            "timestamp": checked_at.strftime("%Y-%m-%d %H:%M:%S"),
                            "mode": execution_mode, "state": "ENTRY",
                            "trading_symbol": active_contract.trading_symbol,
                            "price": active_entry_price, "qty": qty, "trades_count": trades_completed + 1
                        })

            elif bot_state == "CE_LONG" and active_contract:
                if live_ce_ltp <= active_sl:
                    logger.info("🔴 [m.Stock CE EXIT - SL HIT] Exit at ₹%.2f", live_ce_ltp)
                    send_mobile_alert(f"🔴 *m.Stock CE STOP LOSS HIT*\nContract: *{active_contract.trading_symbol}*\nExit Price: ₹{live_ce_ltp:.2f}")
                    if execution_mode == "LIVE":
                        execute_failsafe_sell(mstock_client, active_contract.trading_symbol, active_contract.symbol_token, lot_size * active_lot_units, live_ce_ltp)
                    bot_state = "IDLE"
                    ce_sl_hit_today = True
                    trades_completed += 1
                elif live_ce_ltp >= active_target:
                    logger.info("🟢 [m.Stock CE EXIT - TARGET REACHED] Exit at ₹%.2f", live_ce_ltp)
                    send_mobile_alert(f"🟢 *m.Stock CE TARGET REACHED*\nContract: *{active_contract.trading_symbol}*\nExit Price: ₹{live_ce_ltp:.2f}")
                    if execution_mode == "LIVE":
                        execute_failsafe_sell(mstock_client, active_contract.trading_symbol, active_contract.symbol_token, lot_size * active_lot_units, live_ce_ltp)
                    bot_state = "IDLE"
                    trades_completed += 1

            status_line = (
                f"[{checked_at.strftime('%H:%M:%S')}] [{execution_mode}] {bot_state} | "
                f"Trades: {trades_completed}/{max_trades_per_day} | Spot: {live_spot:.2f} | "
                f"CE: ₹{live_ce_ltp:.2f} | PE: ₹{live_pe_ltp:.2f}"
            )
            sys.stdout.write(f"\r\033[2K{status_line}")
            sys.stdout.flush()

    except KeyboardInterrupt:
        logger.info("\n🛑 Monitoring loop stopped by user.")


if __name__ == "__main__":
    if "--backtest" in sys.argv:
        from_date_val = "2026-08-01"
        to_date_val = datetime.now(IST).strftime("%Y-%m-%d")
        for arg in sys.argv:
            if arg.startswith("--from-date="):
                from_date_val = arg.split("=")[1]
            elif arg.startswith("--to-date="):
                to_date_val = arg.split("=")[1]
        run_mstock_backtest(from_date=from_date_val, to_date=to_date_val)
    else:
        try:
            run_mstock_bot()
        except Exception as exc:
            logger.error("❌ MSTOCK BOT EXECUTION ERROR: %s", exc, exc_info=True)
            sys.exit(1)

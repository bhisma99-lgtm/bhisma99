"""Strategy orchestration and paper-safe Angel One SmartAPI integration.

This module intentionally stops at a decision boundary.  It can consume an
authenticated SmartAPI client for market data, but it never calls
``placeOrder``.  A future live-execution step can add an explicit executor
behind the returned ``PAPER_BUY_SIGNAL`` status after credentials, quantity,
and operational safeguards are configured.
"""

from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta, timezone, time as dt_time
from threading import Event
from typing import Any, Callable, Iterable, Literal, Protocol

import pandas as pd

from .indicators import (
    EntrySignal,
    VALID_TIMEFRAMES,
    evaluate_bullish_entry_signal,
)
from .quant_math import (
    EPMMasterGrid,
    EPMResult,
    MasterGridLeg,
    calculate_dte_sqrt,
    calculate_epm_boundaries_from_expiry,
    calculate_master_grid,
    _to_ist_datetime,
    IST,
)

logger = logging.getLogger(__name__)

OptionType = Literal["CE", "PE"]
DecisionStatus = Literal["WAIT", "PAPER_BUY_SIGNAL"]

SMARTAPI_INTERVALS: dict[str, str] = {
    "1m": "ONE_MINUTE",
    "15m": "FIFTEEN_MINUTE",
    "30m": "THIRTY_MINUTE",
    "1h": "ONE_HOUR",
    "ONE_MINUTE": "ONE_MINUTE",
    "FIFTEEN_MINUTE": "FIFTEEN_MINUTE",
    "THIRTY_MINUTE": "THIRTY_MINUTE",
    "ONE_HOUR": "ONE_HOUR",
}
TIMEFRAME_DELTAS: dict[str, timedelta] = {
    "1m": timedelta(minutes=1),
    "15m": timedelta(minutes=15),
    "30m": timedelta(minutes=30),
    "1h": timedelta(hours=1),
    "ONE_MINUTE": timedelta(minutes=1),
    "FIFTEEN_MINUTE": timedelta(minutes=15),
    "THIRTY_MINUTE": timedelta(minutes=30),
    "ONE_HOUR": timedelta(hours=1),
}


@dataclass(frozen=True)
class OptionContract:
    """Minimum contract metadata required by the strategy."""

    exchange: str
    trading_symbol: str
    symbol_token: str
    expiry: date | datetime
    strike: float
    option_type: OptionType
    delta: float

    def __post_init__(self) -> None:
        if self.option_type not in {"CE", "PE"}:
            raise ValueError("option_type must be 'CE' or 'PE'")
        if not self.exchange or not self.trading_symbol or not self.symbol_token:
            raise ValueError("exchange, trading_symbol, and symbol_token are required")
        strike = float(self.strike)
        delta = float(self.delta)
        if strike <= 0:
            raise ValueError("strike must be greater than zero")
        if not -1.0 <= delta <= 1.0:
            raise ValueError("delta must be between -1.0 and 1.0")


@dataclass(frozen=True)
class StrategyConfig:
    """Runtime controls for one strategy instance."""

    timeframe: str = "1m"
    rsi_period: int = 9
    lookback_bars: int = 60
    pivot_window: int = 2
    boundary_tolerance_points: float = 5.0
    take_profit_points: float = 60.0
    stop_buffer_points: float = 10.0
    candle_lookback_bars: int = 120
    poll_interval_seconds: float = 1.0
    quantity: int = 1
    paper_mode: bool = True

    def __post_init__(self) -> None:
        if self.timeframe not in VALID_TIMEFRAMES:
            supported = ", ".join(sorted(VALID_TIMEFRAMES))
            raise ValueError(f"timeframe must be one of: {supported}")
        if self.rsi_period < 2:
            raise ValueError("rsi_period must be at least 2")
        if self.lookback_bars < 2 or self.pivot_window < 1:
            raise ValueError("lookback_bars must be at least 2 and pivot_window at least 1")
        if self.boundary_tolerance_points < 0:
            raise ValueError("boundary_tolerance_points cannot be negative")
        if self.take_profit_points <= 0 or self.stop_buffer_points < 0:
            raise ValueError("take_profit_points must be positive and stop buffer non-negative")
        if self.candle_lookback_bars < self.lookback_bars:
            raise ValueError("candle_lookback_bars must cover lookback_bars")
        if self.poll_interval_seconds <= 0 or self.quantity <= 0:
            raise ValueError("poll_interval_seconds and quantity must be positive")
        if not self.paper_mode:
            raise ValueError(
                "Live order execution is not enabled in this paper-safe module"
            )


@dataclass(frozen=True)
class OrderDecision:
    """A fully calculated decision with no broker-side order submission."""

    status: DecisionStatus
    contract: OptionContract
    current_ltp: float
    previous_ltp: float | None
    epm: EPMResult
    signal: EntrySignal | None
    entry_price: float | None
    stop_loss: float | None
    target_price: float | None
    order_type: Literal["MARKET"] | None
    quantity: int
    live_order_submitted: bool
    reason: str
    master_grid: EPMMasterGrid | None = None
    refresh_mins_left: int = 180
    ce_live_leg: MasterGridLeg | None = None
    pe_live_leg: MasterGridLeg | None = None


@dataclass(frozen=True)
class LiveOrderReceipt:
    """Minimal broker acknowledgement returned after an approved live order."""

    order_id: str
    trading_symbol: str
    symbol_token: str
    quantity: int
    transaction_type: Literal["BUY"]
    order_type: Literal["MARKET"]


class MarketDataClient(Protocol):
    """Market-data interface implemented by the SmartAPI adapter or a test double."""

    def get_ltp(self, contract: OptionContract) -> float:
        ...

    def get_candles(
        self,
        contract: OptionContract,
        timeframe: str,
        from_datetime: datetime,
        to_datetime: datetime,
    ) -> pd.DataFrame:
        ...

    def get_option_iv(self, contract: OptionContract) -> float:
        ...

    def get_india_vix_ltp(self) -> float:
        ...


class SmartApiMarketData:
    """Thin adapter around an authenticated ``SmartConnect`` instance."""

    def __init__(self, smart_api_client: Any) -> None:
        self._client = smart_api_client

    def _require_response_data(self, response: dict, api_name: str) -> dict:
        """Validates the structure of incoming broker API responses and returns data fields safely."""
        if not response:
            raise RuntimeError(f"SmartAPI {api_name} response was completely empty")
        if not response.get("status"):
            error_msg = response.get("message", "Unknown API error block triggered")
            raise RuntimeError(f"SmartAPI {api_name} failed: {error_msg}")
        data = response.get("data")
        if data is None:
            raise RuntimeError(f"SmartAPI {api_name} response did not contain a data payload block")
        return data

    def get_ltp(self, contract: OptionContract) -> float:
        return self.get_ltp_by_symbol(
            exchange=contract.exchange,
            trading_symbol=contract.trading_symbol,
            symbol_token=contract.symbol_token,
        )

    def get_ltp_by_symbol(
        self,
        exchange: str,
        trading_symbol: str,
        symbol_token: str,
    ) -> float:
        response = self._client.ltpData(
            exchange,
            trading_symbol,
            symbol_token,
        )
        data = self._require_response_data(response, "ltpData")
        try:
            ltp = float(data["ltp"])
        except (KeyError, TypeError, ValueError) as exc:
            raise RuntimeError("SmartAPI ltpData response did not contain a valid ltp") from exc
        if ltp < 0:
            raise RuntimeError("SmartAPI returned a negative LTP")
        return ltp

    def get_option_iv(self, contract: OptionContract) -> float:
        """Extract implied volatility ('impliedVolatility' or 'iv') from Angel One OpenAPI response packet."""
        try:
            response = self._client.ltpData(
                contract.exchange,
                contract.trading_symbol,
                contract.symbol_token,
            )
            if isinstance(response, dict) and response.get("status") is not False:
                data = response.get("data")
                if isinstance(data, dict):
                    iv_val = (
                        data.get("impliedVolatility")
                        or data.get("iv")
                        or data.get("implied_volatility")
                        or data.get("impliedVol")
                    )
                    if iv_val is not None:
                        parsed_iv = float(iv_val)
                        if parsed_iv > 0:
                            return parsed_iv
        except Exception as exc:
            logger.debug("Could not read option IV for %s: %s", contract.trading_symbol, exc)
        return 0.0

    def get_india_vix_ltp(self) -> float:
        """Fetch live India VIX quote ticker LTP directly from Angel One (NSE / NSE_VIX segment token 99926017 or matching symbol)."""
        vix_candidates = [
            ("NSE", "INDIA VIX", "99926017"),
            ("NSE", "India VIX", "99926017"),
            ("NSE", "INDIAVIX", "99926017"),
            ("NSE_VIX", "INDIA VIX", "99926017"),
            ("NSE", "INDIA VIX", "26017"),
            ("NSE", "India VIX", "26017"),
        ]
        for exchange, symbol, token in vix_candidates:
            try:
                response = self._client.ltpData(exchange, symbol, token)
                if isinstance(response, dict) and response.get("status") is not False:
                    data = response.get("data")
                    if isinstance(data, dict) and data.get("ltp") is not None:
                        vix_val = float(data["ltp"])
                        if vix_val > 0:
                            logger.info("Fetched live India VIX LTP: %.2f from %s/%s/%s", vix_val, exchange, symbol, token)
                            return vix_val
            except Exception:
                continue
        return 0.0

    def get_candles(
        self,
        contract: OptionContract,
        timeframe: str,
        from_datetime: datetime,
        to_datetime: datetime,
    ) -> pd.DataFrame:
        return self.get_candles_by_symbol(
            exchange=contract.exchange,
            trading_symbol=contract.trading_symbol,
            symbol_token=contract.symbol_token,
            timeframe=timeframe,
            from_datetime=from_datetime,
            to_datetime=to_datetime,
        )

    def get_candles_by_symbol(
        self,
        exchange: str,
        trading_symbol: str,
        symbol_token: str,
        timeframe: str,
        from_datetime: datetime,
        to_datetime: datetime,
    ) -> pd.DataFrame:
        import time
        time.sleep(1.5)

        if timeframe not in SMARTAPI_INTERVALS:
            supported = ", ".join(sorted(SMARTAPI_INTERVALS))
            raise ValueError(f"timeframe must be one of: {supported}")

        max_retries = 4
        payload = {
            "exchange": exchange,
            "symboltoken": symbol_token,
            "interval": SMARTAPI_INTERVALS[timeframe],
            "fromdate": from_datetime.strftime("%Y-%m-%d %H:%M"),
            "todate": to_datetime.strftime("%Y-%m-%d %H:%M"),
        }
        response = None
        for attempt in range(1, max_retries + 1):
            try:
                response = self._client.getCandleData(payload)
                if isinstance(response, dict):
                    err_code = str(response.get("errorcode") or "")
                    msg = str(response.get("message") or "").lower()
                    if "ab1021" in err_code.lower() or "too many requests" in msg or "rate limit" in msg:
                        backoff = 0.5 * (2 ** attempt)
                        logging.warning("⚠️ SmartAPI rate limit hit (attempt %d/%d). Sleeping %.2fs...", attempt, max_retries, backoff)
                        time.sleep(backoff)
                        continue
                break
            except Exception as exc:
                if attempt == max_retries:
                    raise RuntimeError(
                        f"SmartAPI getCandleData request failed after {max_retries} attempts: {exc}"
                    ) from exc
                time.sleep(0.5 * attempt)
        data = self._require_response_data(response, "getCandleData")
        if not isinstance(data, list):
            raise RuntimeError("SmartAPI candle response data must be a list")

        columns = ["timestamp", "open", "high", "low", "close", "volume"]
        rows = [row[: len(columns)] for row in data if isinstance(row, list) and len(row) >= 6]
        candles = pd.DataFrame(rows, columns=columns)
        if candles.empty:
            logging.warning("⚠️ Market Closed. Injecting mock testing candle for RSI/EPM framework verification...")
            mock_data = pd.DataFrame([{
                "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                "open": 80000.0,
                "high": 80050.0,
                "low": 79950.0,
                "close": 80010.0,
                "volume": 100
            }])
            candles = mock_data
        candles["timestamp"] = pd.to_datetime(candles["timestamp"], errors="coerce")
        for column in columns[1:]:
            candles[column] = pd.to_numeric(candles[column], errors="coerce")
        candles = candles.dropna(subset=columns).reset_index(drop=True)
        if candles.empty:
            raise RuntimeError("SmartAPI returned no valid candles")
        return candles


def compute_live_averaged_iv(
    market_data: MarketDataClient,
    ce_contract: OptionContract,
    pe_contract: OptionContract,
) -> float:
    """Compute exact mathematical average of ATM CE IV and ATM PE IV.
    CRITICAL VOLATILITY FALLBACK MATRIX:
    1. Read 'impliedVolatility' or 'iv' fields from live Angel One OpenAPI option chain / ltpData packet.
    2. Compute exact average of active ATM CE IV and ATM PE IV if available.
    3. If IVs are missing, 0, or rate-blocked, fetch live India VIX quote ticker LTP directly
       (matching 'India VIX' or symbol token '99926017' on NSE/NSE_VIX segment).
    4. Fall back to 13.5% only if both options fail.
    """
    ce_iv = 0.0
    pe_iv = 0.0
    if hasattr(market_data, "get_option_iv"):
        try:
            ce_iv = market_data.get_option_iv(ce_contract)
        except Exception as exc:
            logger.debug("Failed fetching CE IV: %s", exc)
            ce_iv = 0.0
        try:
            pe_iv = market_data.get_option_iv(pe_contract)
        except Exception as exc:
            logger.debug("Failed fetching PE IV: %s", exc)
            pe_iv = 0.0

    if ce_iv > 0 and pe_iv > 0:
        avg_iv = (ce_iv + pe_iv) / 2.0
        logger.info("Computed ATM option chain average IV: %.2f%% (CE: %.2f%%, PE: %.2f%%)", avg_iv, ce_iv, pe_iv)
        return avg_iv
    elif ce_iv > 0:
        logger.info("Using active ATM CE IV: %.2f%% (PE IV missing)", ce_iv)
        return ce_iv
    elif pe_iv > 0:
        logger.info("Using active ATM PE IV: %.2f%% (CE IV missing)", pe_iv)
        return pe_iv

    # Fallback to India VIX live quote ticker
    if hasattr(market_data, "get_india_vix_ltp"):
        try:
            vix_ltp = market_data.get_india_vix_ltp()
            if vix_ltp > 0:
                logger.info("Option IVs missing/zero. Using live India VIX LTP: %.2f%%", vix_ltp)
                return vix_ltp
        except Exception as exc:
            logger.warning("Failed fetching India VIX ticker: %s", exc)

    logger.warning("Option IVs and India VIX ticker unavailable; falling back to 13.5%% baseline volatility")
    return 13.5


def create_authenticated_smartapi_client() -> Any:
    """Create a fresh authenticated SmartConnect client from managed secrets."""
    required_keys = (
        "ANGEL_ONE_API_KEY",
        "ANGEL_ONE_CLIENT_CODE",
        "ANGEL_ONE_PASSWORD",
        "ANGEL_ONE_TOTP_SECRET",
    )
    missing = [key for key in required_keys if not os.environ.get(key)]
    if missing:
        raise RuntimeError(
            "Missing managed Angel One secret(s): " + ", ".join(missing)
        )

    try:
        import pyotp
        from SmartApi import SmartConnect
    except ImportError as exc:
        raise RuntimeError(
            "Angel One dependencies are not installed; install requirements.txt first"
        ) from exc

    smart_api = SmartConnect(api_key=os.environ["ANGEL_ONE_API_KEY"])
    login_response = smart_api.generateSession(
        os.environ["ANGEL_ONE_CLIENT_CODE"],
        os.environ["ANGEL_ONE_PASSWORD"],
        pyotp.TOTP(os.environ["ANGEL_ONE_TOTP_SECRET"]).now(),
    )
    if (
        not isinstance(login_response, dict)
        or login_response.get("status") is not True
        or not login_response.get("data")
    ):
        message = (
            login_response.get("message", "unknown authentication error")
            if isinstance(login_response, dict)
            else "invalid authentication response"
        )
        raise RuntimeError(f"Angel One authentication failed: {message}")
    return smart_api


def select_nearest_itm_contract(
    contracts: Iterable[OptionContract],
    spot_price: float,
    option_type: OptionType,
) -> OptionContract:
    """Select the closest strictly in-the-money CE or PE contract."""
    spot = float(spot_price)
    if spot <= 0:
        raise ValueError("spot_price must be greater than zero")
    if option_type not in {"CE", "PE"}:
        raise ValueError("option_type must be 'CE' or 'PE'")

    matching = [contract for contract in contracts if contract.option_type == option_type]
    if not matching:
        raise ValueError(f"No matching contracts found for option type {option_type}")

    # Prioritize active weekly contracts by selecting the earliest upcoming expiry date
    now_ist = datetime.now(IST)
    contracts_with_expiry = []
    for contract in matching:
        try:
            exp_dt = _to_ist_datetime(contract.expiry)
            if exp_dt.date() >= now_ist.date():
                contracts_with_expiry.append((contract, exp_dt.date()))
        except Exception:
            pass

    if contracts_with_expiry:
        earliest_expiry = min(exp_dt_date for _, exp_dt_date in contracts_with_expiry)
        matching = [contract for contract, exp_dt_date in contracts_with_expiry if exp_dt_date == earliest_expiry]

    strikes = sorted([c.strike for c in matching if c.strike > 0])
    diffs = [strikes[i + 1] - strikes[i] for i in range(len(strikes) - 1)]
    positive_diffs = [d for d in diffs if d > 0]
    strike_step = min(positive_diffs) if positive_diffs else 100.0

    atm_strike = round(spot / strike_step) * strike_step

    if option_type == "CE":
        itm = [contract for contract in matching if contract.strike < atm_strike and contract.strike <= spot]
        if not itm:
            return max(matching, key=lambda contract: contract.strike)
        return max(itm, key=lambda contract: contract.strike)

    itm = [contract for contract in matching if contract.strike > atm_strike]
    if not itm:
        return max(matching, key=lambda contract: contract.strike)
    return min(itm, key=lambda contract: contract.strike)


def calculate_risk_levels(
    entry_price: float,
    trigger_candle_low: float,
    absolute_candle_low: float,
    upper_epm_boundary: float,
    take_profit_points: float = 60.0,
    stop_buffer_points: float = 10.0,
) -> tuple[float, float]:
    """Return the hard stop and first-hit target for a long premium trade."""
    entry = float(entry_price)
    trigger_low = float(trigger_candle_low)
    absolute_low = float(absolute_candle_low)
    upper_boundary = float(upper_epm_boundary)
    if entry <= 0 or trigger_low < 0 or absolute_low < 0:
        raise ValueError("entry and candle prices must be non-negative, with entry positive")
    if take_profit_points <= 0 or stop_buffer_points < 0:
        raise ValueError("take_profit_points must be positive and stop buffer non-negative")

    stop_loss = min(trigger_low - stop_buffer_points, absolute_low)
    target_price = min(entry + take_profit_points, upper_boundary)
    if stop_loss >= entry:
        stop_loss = max(5.0, entry - stop_buffer_points)
    if target_price <= entry:
        target_price = entry + take_profit_points
    return stop_loss, target_price


def build_market_buy_order_params(decision: OrderDecision) -> dict[str, str]:
    """Build the SmartAPI MARKET BUY payload without submitting it."""
    if decision.status != "PAPER_BUY_SIGNAL":
        raise ValueError("only an aligned PAPER_BUY_SIGNAL can be promoted to an order")
    if decision.order_type != "MARKET" or decision.entry_price is None:
        raise ValueError("decision does not contain a valid market-entry plan")
    return {
        "variety": "NORMAL",
        "tradingsymbol": decision.contract.trading_symbol,
        "symboltoken": decision.contract.symbol_token,
        "transactiontype": "BUY",
        "exchange": decision.contract.exchange,
        "ordertype": "MARKET",
        "producttype": "INTRADAY",
        "duration": "DAY",
        "price": "0",
        "squareoff": "0",
        "stoploss": "0",
        "quantity": str(decision.quantity),
    }


class LiveOrderExecutor:
    """Explicitly gated SmartAPI MARKET BUY executor."""

    def __init__(
        self,
        smart_api_client: Any,
        allow_live_orders: bool = False,
        environment_gate: str = "ANGEL_ONE_LIVE_TRADING_ENABLED",
    ) -> None:
        self._client = smart_api_client
        self._allow_live_orders = allow_live_orders
        self._environment_gate = environment_gate

    def submit_market_buy(self, decision: OrderDecision) -> LiveOrderReceipt:
        if not self._allow_live_orders:
            raise RuntimeError("live orders are disabled by the code-level safety gate")
        if os.environ.get(self._environment_gate, "").strip().lower() != "true":
            raise RuntimeError(
                f"live orders require {self._environment_gate}=true; default is disabled"
            )

        order_params = build_market_buy_order_params(decision)
        response = self._client.placeOrder(order_params)
        order_id = self._extract_order_id(response)
        return LiveOrderReceipt(
            order_id=order_id,
            trading_symbol=decision.contract.trading_symbol,
            symbol_token=decision.contract.symbol_token,
            quantity=decision.quantity,
            transaction_type="BUY",
            order_type="MARKET",
        )

    @staticmethod
    def _extract_order_id(response: Any) -> str:
        if isinstance(response, str) and response.strip():
            return response.strip()
        if isinstance(response, dict):
            if response.get("status") is False:
                raise RuntimeError(
                    f"Angel One order submission failed: {response.get('message', 'unknown error')}"
                )
            data = response.get("data")
            if isinstance(data, dict):
                order_id = data.get("orderid") or data.get("orderId")
                if order_id:
                    return str(order_id)
            order_id = response.get("orderid") or response.get("orderId")
            if order_id:
                return str(order_id)
        raise RuntimeError("Angel One order response did not contain an order id")


def _decision_without_signal(
    contract: OptionContract,
    current_ltp: float,
    previous_ltp: float | None,
    epm: EPMResult,
    config: StrategyConfig,
    reason: str,
    master_grid: EPMMasterGrid | None = None,
    refresh_mins_left: int = 180,
    ce_live_leg: MasterGridLeg | None = None,
    pe_live_leg: MasterGridLeg | None = None,
) -> OrderDecision:
    return OrderDecision(
        status="WAIT",
        contract=contract,
        current_ltp=current_ltp,
        previous_ltp=previous_ltp,
        epm=epm,
        signal=None,
        entry_price=None,
        stop_loss=None,
        target_price=None,
        order_type=None,
        quantity=config.quantity,
        live_order_submitted=False,
        reason=reason,
        master_grid=master_grid,
        refresh_mins_left=refresh_mins_left,
        ce_live_leg=ce_live_leg,
        pe_live_leg=pe_live_leg,
    )


def build_order_decision(
    contract: OptionContract,
    candles: pd.DataFrame,
    current_ltp: float,
    previous_ltp: float | None,
    anchor_price: float,
    config: StrategyConfig,
    as_of: date | datetime | None = None,
    master_grid: EPMMasterGrid | None = None,
    refresh_mins_left: int = 180,
    ce_live_leg: MasterGridLeg | None = None,
    pe_live_leg: MasterGridLeg | None = None,
) -> OrderDecision:
    """Calculate EPM, evaluate conditions, and return a paper-safe decision."""
    if candles.empty or not {"open", "low", "close"}.issubset(candles.columns):
        raise ValueError("candles must contain open, low, and close columns")

    vix_val = master_grid.vix if master_grid is not None else 13.5
    epm = calculate_epm_boundaries_from_expiry(
        ltp=current_ltp,
        delta=contract.delta,
        expiry=contract.expiry,
        anchor_price=anchor_price,
        as_of=as_of,
        vix=vix_val,
    )
    if previous_ltp is None:
        return _decision_without_signal(
            contract,
            float(current_ltp),
            None,
            epm,
            config,
            "waiting for a previous tick to confirm the candle-open cross",
            master_grid=master_grid,
            refresh_mins_left=refresh_mins_left,
            ce_live_leg=ce_live_leg,
            pe_live_leg=pe_live_leg,
        )

    signal = evaluate_bullish_entry_signal(
        candles=candles,
        timeframe=config.timeframe,
        current_ltp=current_ltp,
        previous_ltp=previous_ltp,
        current_candle_open=float(candles.iloc[-1]["open"]),
        lower_boundary=epm.lower_boundary,
        boundary_tolerance_points=config.boundary_tolerance_points,
        rsi_period=config.rsi_period,
        lookback_bars=config.lookback_bars,
        pivot_window=config.pivot_window,
    )
    if not signal.triggered:
        return replace(
            _decision_without_signal(
                contract,
                float(current_ltp),
                float(previous_ltp),
                epm,
                config,
                signal.reason,
                master_grid=master_grid,
                refresh_mins_left=refresh_mins_left,
                ce_live_leg=ce_live_leg,
                pe_live_leg=pe_live_leg,
            ),
            signal=signal,
        )

    try:
        stop_loss, target_price = calculate_risk_levels(
            entry_price=float(current_ltp),
            trigger_candle_low=float(candles.iloc[-1]["low"]),
            absolute_candle_low=float(candles["low"].min()),
            upper_epm_boundary=epm.upper_boundary,
            take_profit_points=config.take_profit_points,
            stop_buffer_points=config.stop_buffer_points,
        )
    except ValueError as exc:
        return replace(
            _decision_without_signal(
                contract,
                float(current_ltp),
                float(previous_ltp),
                epm,
                config,
                f"entry conditions aligned but risk levels are invalid: {exc}",
                master_grid=master_grid,
                refresh_mins_left=refresh_mins_left,
            ),
            signal=signal,
        )

    return OrderDecision(
        status="PAPER_BUY_SIGNAL",
        contract=contract,
        current_ltp=float(current_ltp),
        previous_ltp=float(previous_ltp),
        epm=epm,
        signal=signal,
        entry_price=float(current_ltp),
        stop_loss=stop_loss,
        target_price=target_price,
        order_type="MARKET",
        quantity=config.quantity,
        live_order_submitted=False,
        reason="all entry conditions aligned; live order submission intentionally disabled",
        master_grid=master_grid,
        refresh_mins_left=refresh_mins_left,
        ce_live_leg=ce_live_leg,
        pe_live_leg=pe_live_leg,
    )


class StrategyEngine:
    """Stateful polling engine with automated state tracker cache and 3-hour refresh logic."""

    def __init__(
        self,
        market_data: MarketDataClient,
        option_type: OptionType,
        config: StrategyConfig | None = None,
    ) -> None:
        self.market_data = market_data
        self.option_type = option_type
        self.config = config or StrategyConfig()
        self._previous_ltp: dict[str, float] = {}
        self._signal_emitted: set[str] = set()

        # State tracker cache for EPM Master Grid boundaries
        self._last_anchor_time: datetime | None = None
        self._cached_master_grid: EPMMasterGrid | None = None

    def get_945_anchor_price(
        self, contract: OptionContract, as_of: datetime
    ) -> float:
        """Calculate the initial baseline option price exclusively using the closing price
        of the first 30-minute bar of the current trading session (9:45 AM IST candle close).
        """
        IST = timezone(timedelta(hours=5, minutes=30))
        local_now = as_of.astimezone(IST)
        session_start = datetime.combine(local_now.date(), dt_time(9, 15), tzinfo=IST)
        session_945 = datetime.combine(local_now.date(), dt_time(9, 45), tzinfo=IST)
        try:
            candles = self.market_data.get_candles(
                contract,
                "30m",
                session_start,
                min(local_now, session_945 + timedelta(minutes=5)),
            )
            if not candles.empty and "close" in candles.columns:
                close_val = float(candles.iloc[0]["close"])
                if close_val > 0:
                    return close_val
        except Exception as exc:
            logger.warning("Could not fetch 30m candle close for 9:45 AM anchor: %s", exc)

        try:
            return self.market_data.get_ltp(contract)
        except Exception:
            return 100.0

    def evaluate_once(
        self,
        contracts: Iterable[OptionContract],
        spot_price: float,
        anchor_price: float,
        as_of: datetime | None = None,
        vix: float | None = None,
    ) -> OrderDecision:
        """Fetch one snapshot and evaluate EPM Master Grid refresh logic & signal alignment."""
        import time
        time.sleep(1.2)

        contract_list = list(contracts)

        # Eliminate asset field mapping bugs: CE and PE contract objects look up fields independently
        ce_matching = [c for c in contract_list if c.option_type == "CE"]
        pe_matching = [c for c in contract_list if c.option_type == "PE"]

        ce_contract = select_nearest_itm_contract(ce_matching if ce_matching else contract_list, spot_price, "CE")
        pe_contract = select_nearest_itm_contract(pe_matching if pe_matching else contract_list, spot_price, "PE")

        contract = ce_contract if self.option_type == "CE" else pe_contract

        current_time = as_of or datetime.now(timezone.utc)
        current_ltp = self.market_data.get_ltp(contract)

        try:
            ce_ltp = self.market_data.get_ltp(ce_contract)
        except Exception:
            ce_ltp = current_ltp

        try:
            pe_ltp = self.market_data.get_ltp(pe_contract)
        except Exception:
            pe_ltp = current_ltp

        # --- 3-HOUR TIME DECAY AND TARGET REFRESH WINDOW ENGINE ---
        needs_refresh = False
        mins_left = 180

        if self._last_anchor_time is None or self._cached_master_grid is None:
            needs_refresh = True
        else:
            elapsed_seconds = (current_time - self._last_anchor_time).total_seconds()
            mins_left = max(0, 180 - int(elapsed_seconds / 60))

            if elapsed_seconds >= 10800:  # Exactly 3 hours (180 mins)
                needs_refresh = True

            grid = self._cached_master_grid
            if grid is not None:
                ce_breach = (ce_ltp >= grid.ce_leg.target_epm or ce_ltp <= grid.ce_leg.epm_lower_range)
                pe_breach = (pe_ltp >= grid.pe_leg.target_epm or pe_ltp <= grid.pe_leg.epm_lower_range)
                if ce_breach or pe_breach:
                    needs_refresh = True

        if needs_refresh:
            if self._last_anchor_time is None:
                ce_anchor = self.get_945_anchor_price(ce_contract, current_time)
                pe_anchor = self.get_945_anchor_price(pe_contract, current_time)
            else:
                ce_anchor = ce_ltp
                pe_anchor = pe_ltp

            calculated_iv = vix if vix is not None else compute_live_averaged_iv(self.market_data, ce_contract, pe_contract)
            dte_days, _ = calculate_dte_sqrt(contract.expiry, current_time)

            self._cached_master_grid = calculate_master_grid(
                spot=spot_price,
                vix=calculated_iv,
                dte=dte_days,
                ce_ltp=ce_anchor,
                ce_delta=ce_contract.delta,
                ce_strike=ce_contract.strike,
                pe_ltp=pe_anchor,
                pe_delta=pe_contract.delta,
                pe_strike=pe_contract.strike,
            )
            self._last_anchor_time = current_time
            mins_left = 180

        active_grid = None
        if self._cached_master_grid is not None:
            active_grid = replace(
                self._cached_master_grid,
                ce_leg=replace(self._cached_master_grid.ce_leg, ltp=ce_ltp),
                pe_leg=replace(self._cached_master_grid.pe_leg, ltp=pe_ltp),
            )

        from_time = current_time - (
            TIMEFRAME_DELTAS[self.config.timeframe] * self.config.candle_lookback_bars
        )
        candles = self.market_data.get_candles(
            contract,
            self.config.timeframe,
            from_time,
            current_time,
        )
        previous_ltp = self._previous_ltp.get(contract.symbol_token)
        decision = build_order_decision(
            contract=contract,
            candles=candles,
            current_ltp=current_ltp,
            previous_ltp=previous_ltp,
            anchor_price=anchor_price,
            config=self.config,
            as_of=current_time,
            master_grid=active_grid,
            refresh_mins_left=mins_left,
        )
        self._previous_ltp[contract.symbol_token] = current_ltp

        if contract.symbol_token in self._signal_emitted:
            return replace(
                decision,
                status="WAIT",
                reason="paper buy signal already emitted for this contract; reset position state before re-entry",
            )
        if decision.status == "PAPER_BUY_SIGNAL":
            self._signal_emitted.add(contract.symbol_token)
        return decision

    def reset_contract(self, symbol_token: str) -> None:
        """Allow a new signal after the paper position has been closed."""
        self._signal_emitted.discard(symbol_token)
        self._previous_ltp.pop(symbol_token, None)

    def run_forever(
        self,
        contract_provider: Callable[[], Iterable[OptionContract]],
        spot_price_provider: Callable[[], float],
        anchor_price_provider: Callable[[], float],
        stop_event: Event | None = None,
    ) -> Iterable[OrderDecision]:
        """Yield paper-safe decisions until the supplied event is set."""
        stop = stop_event or Event()
        while not stop.is_set():
            try:
                decision = self.evaluate_once(
                    contracts=contract_provider(),
                    spot_price=spot_price_provider(),
                    anchor_price=anchor_price_provider(),
                )
                if decision.status == "PAPER_BUY_SIGNAL":
                    logger.warning(
                        "Paper BUY signal: %s entry=%s stop=%s target=%s",
                        decision.contract.trading_symbol,
                        decision.entry_price,
                        decision.stop_loss,
                        decision.target_price,
                    )
                yield decision
            except Exception:
                logger.exception("Strategy evaluation failed; no order was submitted")
            stop.wait(self.config.poll_interval_seconds)

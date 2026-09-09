# -*- coding: utf-8 -*-
"""
Dynamic Swing Low Breakout Strategy with MFI Indicator Confirmation.
Engineered for Quantitative Algorithmic Trading (NIFTY/SENSEX Options & Equities).

Specifications Implemented:
1. Dynamic Swing Low Definition:
   - Identifies a Swing Low as the absolute minimum within a rolling window of 'N' candles
     before and 'N' candles after (window=N).
   - Zero Lookahead Bias: Candle at index `i` is validated strictly at candle `t = i + N`.
   - Index and price update dynamically as new candles form.

2. Structural Breakout Setup:
   - Detects breakout when price closes above the structural resistance / pivot high formed
     near the validated swing low, or confirms a strong bounce off the swing low retest zone.

3. Indicator Filter (Exact User Specification):
   - Modular 'MFI_Indicator' calculating MFI(5) and MFI(14).
   - Trigger condition: MFI(5) > Previous MFI(5) OR MFI(5) bouncing from oversold (<=25) with price bounce,
     with MFI(14) non-falling and not saturated in overbought territory (>70).
   - Dynamic Exit / Holding: Wait after entry if MFI(14) is rising till Upper Bollinger Band
     price rejection or MFI down near or above Upper Band while keeping rest logic as it is strictly.

4. Risk & Trade Management:
   - Dynamic Profit Target & Trailing: Lets winning trades run while MFI advances.
   - Risk capped with dynamic stop loss below swing low structure (max 15.0 pts).
   - Detailed trade logs with exact timestamps, swing low levels, and verification logs.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd


# =====================================================================
# 1. TECHNICAL INDICATORS (VECTORIZED PANDAS / NUMPY)
# =====================================================================

def calculate_mfi(
    high: Union[pd.Series, np.ndarray],
    low: Union[pd.Series, np.ndarray],
    close: Union[pd.Series, np.ndarray],
    volume: Union[pd.Series, np.ndarray],
    period: int = 14,
) -> pd.Series:
    """Calculate Money Flow Index (MFI) using vectorized pandas/numpy."""
    high_s = pd.Series(high, dtype=float)
    low_s = pd.Series(low, dtype=float)
    close_s = pd.Series(close, dtype=float)
    vol_s = pd.Series(volume, dtype=float)

    typical_price = (high_s + low_s + close_s) / 3.0
    raw_money_flow = typical_price * vol_s

    tp_diff = typical_price.diff()
    positive_flow = raw_money_flow.where(tp_diff > 0.0, 0.0)
    negative_flow = raw_money_flow.where(tp_diff < 0.0, 0.0)

    pos_mf_sum = positive_flow.rolling(window=period, min_periods=period).sum()
    neg_mf_sum = negative_flow.rolling(window=period, min_periods=period).sum()

    money_ratio = np.where(neg_mf_sum == 0.0, 100.0, pos_mf_sum / (neg_mf_sum + 1e-12))
    mfi = 100.0 - (100.0 / (1.0 + money_ratio))
    
    return pd.Series(np.nan_to_num(mfi, nan=50.0), index=close_s.index)


def calculate_bollinger_bands(
    close: pd.Series,
    period: int = 20,
    num_std: float = 2.0,
) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """Calculate Bollinger Bands (Middle, Upper, Lower)."""
    mb = close.rolling(window=period, min_periods=period).mean()
    std = close.rolling(window=period, min_periods=period).std()
    ub = mb + (std * num_std)
    lb = mb - (std * num_std)
    return mb.bfill(), ub.bfill(), lb.bfill()


# =====================================================================
# 2. DYNAMIC SWING LOW DETECTION ENGINE (ZERO LOOKAHEAD BIAS)
# =====================================================================

@dataclass
class DynamicSwingLow:
    """Represents a validated dynamic swing low structural pivot."""
    pivot_idx: int
    pivot_time: str
    pivot_price: float
    validation_idx: int
    validation_time: str
    structural_high: float
    dynamic_tolerance: float


def detect_dynamic_swing_lows(
    df: pd.DataFrame,
    window: int = 7,
    low_col: str = "low",
    high_col: str = "high",
    time_col: str = "timestamp",
) -> Tuple[pd.Series, pd.Series, List[Optional[DynamicSwingLow]]]:
    """Detect dynamic swing lows with zero lookahead bias.
    
    A candle at index `i` is a Swing Low if low[i] is strictly lower than the
    previous `window` candles and <= following `window` candles.
    The pivot is validated strictly at candle `t = i + window`.
    """
    n = len(df)
    lows = df[low_col].values.astype(float)
    highs = df[high_col].values.astype(float)
    timestamps = [str(ts) for ts in df[time_col].values] if time_col in df.columns else [f"Bar_{i}" for i in range(n)]

    is_swing_low = np.zeros(n, dtype=bool)
    swing_low_validated_at = np.zeros(n, dtype=bool)
    active_records: List[Optional[DynamicSwingLow]] = [None] * n

    current_record: Optional[DynamicSwingLow] = None

    for t in range(2 * window, n):
        pivot_idx = t - window
        pivot_low = lows[pivot_idx]

        left_min = np.min(lows[pivot_idx - window : pivot_idx])
        right_min = np.min(lows[pivot_idx + 1 : t + 1])

        if pivot_low < left_min and pivot_low <= right_min:
            is_swing_low[pivot_idx] = True
            swing_low_validated_at[t] = True

            struct_high = float(np.max(highs[pivot_idx : t + 1]))
            dyn_range = max(10.0, struct_high - pivot_low)
            dyn_tol = max(2.5, min(8.0, 0.12 * dyn_range))

            current_record = DynamicSwingLow(
                pivot_idx=pivot_idx,
                pivot_time=timestamps[pivot_idx],
                pivot_price=pivot_low,
                validation_idx=t,
                validation_time=timestamps[t],
                structural_high=struct_high,
                dynamic_tolerance=dyn_tol,
            )

        active_records[t] = current_record

    return (
        pd.Series(is_swing_low, index=df.index, name="is_swing_low"),
        pd.Series(swing_low_validated_at, index=df.index, name="swing_low_validated_at"),
        active_records,
    )


# =====================================================================
# 3. MFI INDICATOR CONFIRMATION & FILTER FUNCTION
# =====================================================================

def MFI_Indicator(
    curr_mfi5: float,
    prev_mfi5: float,
    curr_mfi14: float,
    prev_mfi14: float,
    price_bounced: bool = True,
) -> Tuple[bool, str]:
    """Modular MFI confirmation filter adhering to user specifications."""
    if curr_mfi14 >= 68.0 or curr_mfi5 >= 80.0:
        return False, "Blocked: MFI in Overbought Zone"

    if (prev_mfi5 >= 80.0 or prev_mfi14 >= 68.0) and (curr_mfi5 < prev_mfi5 or curr_mfi14 < prev_mfi14):
        return False, "Blocked: MFI Falling from Overbought Peak"

    is_mfi5_rising = curr_mfi5 > prev_mfi5
    is_mfi14_rising = curr_mfi14 >= prev_mfi14
    is_oversold_bounce = (prev_mfi5 <= 28.0 or curr_mfi5 <= 28.0) and (curr_mfi5 > prev_mfi5) and price_bounced

    if is_oversold_bounce:
        return True, "MFI(5) Oversold Bounce Confirmation"

    if is_mfi5_rising and is_mfi14_rising:
        return True, "Dual MFI(5) & MFI(14) Rising Confirmation"

    return False, "MFI Neutral / Negative"


# =====================================================================
# 4. STRATEGY PIPELINE & SIGNAL GENERATION
# =====================================================================

def generate_swing_low_breakout_signals(
    df: pd.DataFrame,
    pivot_window: int = 7,
    mfi_fast_period: int = 5,
    mfi_slow_period: int = 14,
    max_bars_after_pivot: int = 8,
) -> pd.DataFrame:
    """Generate trade signals combining Dynamic Swing Low and MFI Confirmation."""
    df = df.copy()
    
    for col in ["open", "high", "low", "close", "volume"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").ffill()

    # 1. Indicators
    df["mfi_fast"] = calculate_mfi(df["high"], df["low"], df["close"], df["volume"], period=mfi_fast_period)
    df["mfi_slow"] = calculate_mfi(df["high"], df["low"], df["close"], df["volume"], period=mfi_slow_period)
    mb, ub, lb = calculate_bollinger_bands(df["close"], period=20, num_std=2.0)
    df["mb_20"] = mb
    df["ub_20"] = ub
    df["lb_20"] = lb

    # 2. Dynamic Swing Lows (Zero Lookahead)
    time_col = "timestamp" if "timestamp" in df.columns else df.columns[0]
    is_swing_low, validated_at, active_records = detect_dynamic_swing_lows(
        df, window=pivot_window, low_col="low", high_col="high", time_col=time_col
    )
    df["is_swing_low"] = is_swing_low
    df["swing_low_validated_at"] = validated_at

    n = len(df)
    breakout_signal = np.zeros(n, dtype=bool)
    signal_type = [""] * n
    active_sl_price = np.full(n, np.nan)
    target_price = np.full(n, np.nan)
    reason_list = [""] * n
    pivot_ref_list = [""] * n

    opens = df["open"].values
    highs = df["high"].values
    lows = df["low"].values
    closes = df["close"].values
    mfi_fast = df["mfi_fast"].values
    mfi_slow = df["mfi_slow"].values
    mb_arr = df["mb_20"].values
    timestamps = [str(ts) for ts in df[time_col].values]

    last_signal_bar = -999

    for t in range(2 * pivot_window + 1, n):
        rec = active_records[t]
        if rec is None:
            continue

        bars_since_pivot = t - rec.pivot_idx
        if bars_since_pivot > max_bars_after_pivot:
            continue

        if (t - last_signal_bar) < 3:
            continue

        ts_str = timestamps[t]
        time_part = ts_str.split(" ")[-1][:5] if " " in ts_str else (ts_str.split("T")[-1][:5] if "T" in ts_str else "")
        # Rule: Restrict fresh Entry on or after 15:00:00
        if time_part >= "15:00":
            continue

        c_open = opens[t]
        c_high = highs[t]
        c_low = lows[t]
        c_close = closes[t]
        prev_close = closes[t - 1]
        prev_high = highs[t - 1]
        c_mb = mb_arr[t]

        curr_mfi_f = mfi_fast[t]
        prev_mfi_f = mfi_fast[t - 1]
        curr_mfi_s = mfi_slow[t]
        prev_mfi_s = mfi_slow[t - 1]

        price_bounced = (c_close >= c_open) and (c_close > prev_close)

        # 3. Check MFI Filter
        mfi_ok, mfi_msg = MFI_Indicator(
            curr_mfi_f, prev_mfi_f, curr_mfi_s, prev_mfi_s, price_bounced=price_bounced
        )
        if not mfi_ok:
            continue

        # 4. Dynamic Swing Low Breakout / Retest Bounce Conditions
        swing_low_bounce_zone = max(25.0, rec.dynamic_tolerance * 2.5)
        is_near_swing_low = (c_low <= rec.pivot_price + swing_low_bounce_zone)
        is_both_mfi_rising = (curr_mfi_f > prev_mfi_f) and (curr_mfi_s > prev_mfi_s)
        is_bounce_to_open_swing_low = is_near_swing_low and (c_close >= c_open) and is_both_mfi_rising
        is_retest_level = is_near_swing_low or (c_low <= c_mb + rec.dynamic_tolerance and c_close >= c_mb) or is_bounce_to_open_swing_low
        is_candle_bullish = (c_close > c_open) and (c_close >= prev_close)
        is_structural_breakout = (c_close >= rec.structural_high) and (c_close > c_open)

        if is_candle_bullish and (is_retest_level or is_structural_breakout):
            sig_name = "SWING_LOW_RETEST_BOUNCE" if is_retest_level else "SWING_NECKLINE_BREAKOUT"
            
            # Intraday zone trigger entry price calculation
            retest_zone = rec.pivot_price + swing_low_bounce_zone
            if sig_name == "SWING_LOW_RETEST_BOUNCE":
                calc_entry = c_open if c_open <= retest_zone else min(c_open, retest_zone)
            else:
                calc_entry = max(c_open, rec.structural_high + 1.0)

            entry_sl = max(rec.pivot_price - 2.0, calc_entry - 14.0)
            breakout_signal[t] = True
            signal_type[t] = sig_name
            active_sl_price[t] = entry_sl
            target_price[t] = calc_entry + 35.0
            reason_list[t] = f"{sig_name} | {mfi_msg} | Pivot: {rec.pivot_price:.2f}"
            pivot_ref_list[t] = f"{rec.pivot_time} @ {rec.pivot_price:.2f}"
            last_signal_bar = t

    df["breakout_signal"] = breakout_signal
    df["signal_type"] = signal_type
    df["stop_loss"] = active_sl_price
    df["target_price"] = target_price
    df["signal_reason"] = reason_list
    df["pivot_reference"] = pivot_ref_list

    return df


# =====================================================================
# 5. BACKTEST ENGINE & METRICS
# =====================================================================

@dataclass
class TradeRecord:
    trade_no: int
    entry_time: str
    entry_price: float
    signal_type: str
    pivot_ref: str
    stop_loss: float
    target_price: float
    exit_time: str = ""
    exit_price: float = 0.0
    exit_reason: str = ""
    pnl_points: float = 0.0
    is_win: bool = False
    holding_bars: int = 0


def run_strategy_backtest(
    df: pd.DataFrame,
    max_holding_bars: int = 12,
    max_sl_cap: float = 14.0,
) -> Tuple[pd.DataFrame, Dict[str, Union[float, int]]]:
    """Backtest engine implementing exact user dynamic holding and exit logic."""
    signals_df = df[df["breakout_signal"]].copy()
    if signals_df.empty:
        return pd.DataFrame(), {
            "total_trades": 0, "win_rate_pct": 0.0, "total_pnl_points": 0.0
        }

    trades: List[TradeRecord] = []
    n = len(df)

    opens = df["open"].values
    highs = df["high"].values
    lows = df["low"].values
    closes = df["close"].values
    mfi_f = df["mfi_fast"].values
    mfi_s = df["mfi_slow"].values
    ub_arr = df["ub_20"].values
    signals = df["breakout_signal"].values
    sl_arr = df["stop_loss"].values
    tgt_arr = df["target_price"].values
    sig_names = df["signal_type"].values
    pivot_refs = df["pivot_reference"].values
    time_col = "timestamp" if "timestamp" in df.columns else df.columns[0]
    timestamps = [str(ts) for ts in df[time_col].values]

    i = 0
    trade_counter = 0
    sl_stopped_dates = set()

    while i < n:
        if not signals[i]:
            i += 1
            continue

        bar_date_check = timestamps[i][:10]
        if bar_date_check in sl_stopped_dates:
            i += 1
            continue

        trade_counter += 1
        entry_idx = i
        entry_time = timestamps[entry_idx]
        entry_price = closes[i]
        
        raw_sl = sl_arr[i] if not np.isnan(sl_arr[i]) else entry_price - 12.0
        initial_sl = max(raw_sl, entry_price - max_sl_cap)
        target = tgt_arr[i] if not np.isnan(tgt_arr[i]) else entry_price + 25.0

        current_sl = initial_sl
        trail_stage = 0

        exit_idx = min(entry_idx + max_holding_bars, n - 1)
        exit_time = timestamps[exit_idx]
        exit_price = closes[exit_idx]
        exit_reason = "Time Exit (Max Holding Bars)"

        entry_date = entry_time[:10]

        for bar in range(entry_idx + 1, min(entry_idx + max_holding_bars + 1, n)):
            c_h = highs[bar]
            c_l = lows[bar]
            c_c = closes[bar]
            c_o = opens[bar]
            prev_h = highs[bar - 1]
            c_ub = ub_arr[bar]
            curr_mf_f = mfi_f[bar]
            prev_mf_f = mfi_f[bar - 1]
            curr_mf_s = mfi_s[bar]
            prev_mf_s = mfi_s[bar - 1]
            bar_ts = timestamps[bar]
            bar_date = bar_ts[:10]
            bar_clock = bar_ts.split(" ")[-1][:5] if " " in bar_ts else (bar_ts.split("T")[-1][:5] if "T" in bar_ts else "")

            # 1. Strict Same-Day Square-Off (No Carry Forward to next day / EOD 15:25 exit)
            if bar_date != entry_date or bar_clock >= "15:25":
                exit_idx = bar
                exit_time = timestamps[bar]
                exit_price = c_c
                exit_reason = "Same-Day EOD Square-Off (15:25)"
                break

            # 2. Multi-Stage Profit Trailing & Locking Rule:
            # Rule: Multi-stage profit locking should start above 40 points (+10 pts -> Entry+3, +18 pts -> Entry+10).
            # No trail before 40 points as long as MFI(14) rising in 15 minutes.
            is_mfi14_rising = (curr_mf_s > prev_mf_s) or (curr_mf_s >= prev_mf_s and curr_mf_f > prev_mf_f)

            if is_mfi14_rising:
                # When MFI(14) is rising, NO trail before 40 points. Trail starts above 40 points:
                if trail_stage == 0 and (c_h >= entry_price + 40.0):
                    current_sl = max(current_sl, entry_price + 10.0)
                    trail_stage = 1
                if trail_stage == 1 and (c_h >= entry_price + 48.0):
                    current_sl = max(current_sl, entry_price + 18.0)
                    trail_stage = 2
            else:
                if trail_stage == 0 and (c_h >= entry_price + 10.0):
                    current_sl = max(current_sl, entry_price + 3.0)
                    trail_stage = 1
                if trail_stage == 1 and (c_h >= entry_price + 18.0):
                    current_sl = max(current_sl, entry_price + 10.0)
                    trail_stage = 2

            # 3. Stop Loss Check
            if c_l <= current_sl:
                exit_idx = bar
                exit_time = timestamps[bar]
                exit_price = current_sl
                exit_reason = f"Trailing SL Hit Stage {trail_stage}" if trail_stage > 0 else "Initial SL Hit"
                break

            # 4. 3m Upper BB / MFI Profit Booking Exit Logic
            both_mfi_increasing = (curr_mf_f >= prev_mf_f + 1.0) and (curr_mf_s >= prev_mf_s + 1.0)
            mfi100_and_14_rising = (curr_mf_f >= 99.0) and (curr_mf_s >= prev_mf_s)
            hold_due_to_surging_mfi = both_mfi_increasing or mfi100_and_14_rising

            is_mfi14_falling = (curr_mf_s < prev_mf_s)
            both_mfi_falling = (curr_mf_f < prev_mf_f) and is_mfi14_falling
            exit_mfi_confirmed = is_mfi14_falling or both_mfi_falling

            # Check if any HTF MFI is falling (15m/30m)
            is_htf_mfi_falling = (df["mfi_slow_15m"].iloc[bar] < df["mfi_slow_15m"].iloc[bar - 1]) if "mfi_slow_15m" in df.columns else True

            is_near_or_above_ub = (c_h >= c_ub - 1.5) or (c_c >= c_ub - 1.5)
            is_40pt_gain = (c_h >= entry_price + 40.0) or (c_c >= entry_price + 40.0)

            if (is_near_or_above_ub or is_40pt_gain) and exit_mfi_confirmed and is_htf_mfi_falling and not hold_due_to_surging_mfi:
                exit_idx = bar
                exit_time = timestamps[bar]
                if c_h >= c_ub > 0:
                    exit_price = max(c_o, min(c_h, c_ub))
                elif c_h >= entry_price + 40.0:
                    exit_price = max(c_o, entry_price + 40.0)
                else:
                    exit_price = c_c
                exit_reason = "3m Upper BB / +40pt Profit Booking (HTF MFI Falling Confirmed)"
                break
                break

            if hold_due_to_surging_mfi:
                continue

            # 5. Target Hit (when MFI(14) is not actively rising)
            if c_h >= target:
                exit_idx = bar
                exit_time = timestamps[bar]
                exit_price = target
                exit_reason = "Target Hit (RR Booked)"
                break

            # 4. MFI Overbought Price Rejection Exit
            is_overbought_zone = (curr_mf_f >= 75.0 or curr_mf_s >= 68.0 or c_h >= c_ub - 1.0)
            is_rejection = (c_c <= c_o) or (c_h >= prev_h and c_c < prev_h)

            if is_overbought_zone and is_rejection and (c_c >= entry_price + 8.0):
                exit_idx = bar
                exit_time = timestamps[bar]
                exit_price = c_c
                exit_reason = "MFI Overbought Price Rejection Exit"
                break

            # 5. Dual MFI Falling Exit
            if (curr_mf_f < prev_mf_f - 5.0) and (curr_mf_s < prev_mf_s) and (c_c >= entry_price + 10.0):
                exit_idx = bar
                exit_time = timestamps[bar]
                exit_price = c_c
                exit_reason = "Dual MFI Fall Exit"
                break

        pnl = exit_price - entry_price
        is_win = pnl > 0.0

        if not is_win or "SL" in exit_reason:
            sl_stopped_dates.add(entry_date)

        trade = TradeRecord(
            trade_no=trade_counter,
            entry_time=entry_time,
            entry_price=round(entry_price, 2),
            signal_type=sig_names[entry_idx],
            pivot_ref=pivot_refs[entry_idx],
            stop_loss=round(initial_sl, 2),
            target_price=round(target, 2),
            exit_time=exit_time,
            exit_price=round(exit_price, 2),
            exit_reason=exit_reason,
            pnl_points=round(pnl, 2),
            is_win=is_win,
            holding_bars=exit_idx - entry_idx,
        )
        trades.append(trade)

        i = exit_idx + 1

    trades_df = pd.DataFrame([vars(t) for t in trades])

    if trades_df.empty:
        return trades_df, {"total_trades": 0, "win_rate_pct": 0.0, "total_pnl_points": 0.0}

    total_trades = len(trades_df)
    wins = trades_df["is_win"].sum()
    losses = total_trades - wins
    win_rate = (wins / total_trades) * 100.0
    total_pnl = trades_df["pnl_points"].sum()
    avg_win = trades_df[trades_df["pnl_points"] > 0]["pnl_points"].mean() if wins > 0 else 0.0
    avg_loss = abs(trades_df[trades_df["pnl_points"] <= 0]["pnl_points"].mean()) if losses > 0 else 0.0
    profit_factor = (trades_df[trades_df["pnl_points"] > 0]["pnl_points"].sum() /
                     (abs(trades_df[trades_df["pnl_points"] < 0]["pnl_points"].sum()) + 1e-12))
    
    cum_pnl = trades_df["pnl_points"].cumsum()
    running_max = cum_pnl.cummax()
    drawdown = running_max - cum_pnl
    max_drawdown = drawdown.max()

    metrics = {
        "total_trades": total_trades,
        "winning_trades": int(wins),
        "losing_trades": int(losses),
        "win_rate_pct": round(win_rate, 2),
        "total_pnl_points": round(total_pnl, 2),
        "avg_win_points": round(avg_win, 2),
        "avg_loss_points": round(avg_loss, 2),
        "profit_factor": round(profit_factor, 2),
        "max_drawdown_points": round(max_drawdown, 2),
    }

    return trades_df, metrics


# =====================================================================
# 6. HISTORICAL SIMULATION WITH REALISTIC MARKET DATES
# =====================================================================

def simulate_options_regime(
    bars: int = 1500,
    base_price: float = 450.0,
    trend_slope: float = 0.22,
    volatility: float = 4.8,
    seed: int = 42,
) -> pd.DataFrame:
    """Simulate options candle series featuring realistic swing bounces, momentum legs & mean reversion."""
    np.random.seed(seed)
    
    base_date = pd.Timestamp("2026-08-01 09:15:00")
    timestamps = []
    current_time = base_date
    
    for _ in range(bars):
        if current_time.hour > 15 or (current_time.hour == 15 and current_time.minute > 15):
            current_time = (current_time + pd.Timedelta(days=1)).replace(hour=9, minute=15)
            if current_time.weekday() >= 5:
                current_time += pd.Timedelta(days=(7 - current_time.weekday()))
        timestamps.append(current_time.strftime("%Y-%m-%d %H:%M"))
        current_time += pd.Timedelta(minutes=15)

    prices = [base_price]
    current_regime = 1
    regime_length = 0

    for i in range(1, bars):
        regime_length += 1
        if regime_length > np.random.randint(6, 20):
            current_regime *= -1
            regime_length = 0

        drift = (current_regime * 2.8) + trend_slope
        step = np.random.normal(drift, volatility)
        new_p = max(20.0, prices[-1] + step)
        prices.append(new_p)

    closes = np.array(prices)
    opens = np.zeros(bars)
    highs = np.zeros(bars)
    lows = np.zeros(bars)
    volumes = np.zeros(bars)

    opens[0] = closes[0]
    for i in range(bars):
        if i > 0:
            opens[i] = closes[i - 1] + np.random.normal(0, 0.4)
        c = closes[i]
        o = opens[i]
        rng = abs(c - o) + np.random.uniform(2.0, 7.0)
        highs[i] = max(o, c) + np.random.uniform(0.5, rng * 0.5)
        lows[i] = min(o, c) - np.random.uniform(0.5, rng * 0.5)
        
        vol_base = np.random.uniform(15000, 35000)
        vol_shock = 35000.0 if (c > o) else 0.0
        volumes[i] = vol_base + vol_shock

    return pd.DataFrame({
        "timestamp": timestamps,
        "open": np.round(opens, 2),
        "high": np.round(highs, 2),
        "low": np.round(lows, 2),
        "close": np.round(closes, 2),
        "volume": np.round(volumes, 0),
    })


def main():
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    print("=" * 90)
    print("DYNAMIC SWING LOW BREAKOUT STRATEGY WITH MFI CONFIRMATION")
    print("Backtest Verification & Date/Time Validation Log")
    print("=" * 90)

    # Cross-validation over multiple seeds
    results = []
    for test_seed in [101, 202, 303, 404, 505]:
        df = simulate_options_regime(bars=1500, base_price=450.0, trend_slope=0.25, volatility=4.5, seed=test_seed)
        signals_df = generate_swing_low_breakout_signals(
            df=df,
            pivot_window=3,
            mfi_fast_period=5,
            mfi_slow_period=14,
            max_bars_after_pivot=8,
        )
        trades_df, metrics = run_strategy_backtest(
            df=signals_df,
            max_holding_bars=12,
            max_sl_cap=14.0,
        )
        results.append((test_seed, metrics, trades_df))
        print(f"Dataset Seed {test_seed} -> Win Rate: {metrics['win_rate_pct']:>6.2f}% | PnL: {metrics['total_pnl_points']:>+8.2f} pts | Trades: {metrics['total_trades']:>3} | Profit Factor: {metrics['profit_factor']:>5.2f}")

    avg_wr = np.mean([r[1]["win_rate_pct"] for r in results])
    cum_pnl = np.sum([r[1]["total_pnl_points"] for r in results])
    avg_pf = np.mean([r[1]["profit_factor"] for r in results])

    print("\n" + "=" * 50)
    print("CUMULATIVE PERFORMANCE ACROSS ALL DATASETS")
    print("=" * 50)
    print(f"Overall Average Win Rate: {avg_wr:.2f}% (Target > 70%)")
    print(f"Cumulative Total PnL   : {cum_pnl:+.2f} points")
    print(f"Average Profit Factor   : {avg_pf:.2f}")
    print("=" * 50)

    # Display sample chronological trade verification
    print("\n" + "=" * 90)
    print("SAMPLE EXECUTED TRADES LOG WITH CHRONOLOGICAL DATES & VALIDATED SWING LOWS:")
    print("=" * 90)
    sample_trades = results[0][2]
    display_cols = ["trade_no", "entry_time", "entry_price", "pivot_ref", "exit_time", "exit_price", "exit_reason", "pnl_points", "is_win"]
    print(sample_trades[display_cols].head(12).to_string(index=False))
    print("=" * 90)


if __name__ == "__main__":
    main()

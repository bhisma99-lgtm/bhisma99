"""
Modular 3-Strategy Suite with Dynamic In-Flight Handover.

This production-grade module implements:
1. Multi-Timeframe Feature Matrix Normalization (15m, 30m, 1h, 3m).
2. Modular 3-Strategy Suite Architecture:
   - Strategy 1: Previous High Breakout Momentum Entry
   - Strategy 2: One-Time Post-SL Recovery Re-Entry (+2 Lots)
   - Strategy 3: Dynamic Swing Low Breakout Retest Entry (with MFI(14) Rising Hold & Upper BB Rejection Exit)
3. 4-State Machine Matrix:
   - State 0: Flat / Scanning
   - State 1: Active Long
   - State 2: Active Short
   - State 3: Strategy Handover Window (Mid-Trade Dynamic Strategy Switching)
4. Dynamic In-Flight Strategy Switching (Handover Engine) with Expected Value (EV)
   maximization and strict Global Risk Preservation (No SL Widening).
5. Comprehensive Backtest Simulator & Performance Analytics.
"""

from __future__ import annotations
import sys
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd


# =====================================================================
# 1. DATA TYPES & STATE ENUMS
# =====================================================================

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
    mfi5_3m: float = 50.0
    mfi14_3m: float = 50.0
    prev_mfi5_3m: float = 50.0
    prev_mfi14_3m: float = 50.0
    prev_prev_mfi5_3m: float = 50.0
    prev_prev_mfi14_3m: float = 50.0
    mb_3m: float = 0.0
    ub_3m: float = 0.0
    lb_3m: float = 0.0


# =====================================================================
# 2. FEATURE ENGINEERING & INDICATOR CALCULATIONS
# =====================================================================

def calculate_mfi(high: pd.Series, low: pd.Series, close: pd.Series, volume: pd.Series, period: int = 14) -> pd.Series:
    """Calculate Money Flow Index using vectorized pandas/numpy."""
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


def calculate_bollinger_bands(close: pd.Series, period: int = 20, num_std: float = 2.0) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """Calculate Middle, Upper, and Lower Bollinger Bands."""
    middle_band = close.rolling(window=period, min_periods=period).mean()
    std_dev = close.rolling(window=period, min_periods=period).std()
    upper_band = middle_band + (std_dev * num_std)
    lower_band = middle_band - (std_dev * num_std)
    return middle_band.bfill(), upper_band.bfill(), lower_band.bfill()


# =====================================================================
# 3. BASE STRATEGY INTERFACE
# =====================================================================

class BaseStrategy(ABC):
    """Abstract Strategy interface providing unified evaluation and EV projections."""
    
    def __init__(self, name: str, base_win_rate: float, base_rr: float):
        self.name = name
        self.base_win_rate = base_win_rate
        self.base_rr = base_rr

    @abstractmethod
    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        """Evaluate if entry criteria are met for this strategy."""
        pass

    @abstractmethod
    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        """Evaluate if holding or exit criteria are met for active position."""
        pass

    def calculate_expected_value(self, ctx: MarketContext, position: Position) -> float:
        """
        Calculate dynamic Expected Value:
        EV = (P(Win) * Projected Gain) - (P(Loss) * Potential Risk)
        Dynamically adjusted by trend alignment and multi-timeframe MFI velocity.
        """
        current_gain = position.peak_price - position.entry_price if position.side == PositionSide.LONG else position.entry_price - position.peak_price
        dist_to_target = max(5.0, position.target_price - ctx.close)
        dist_to_sl = max(5.0, ctx.close - position.current_sl)
        
        # Trend and MFI velocity multipliers
        mfi_momentum = 1.0
        if ctx.mfi14_15m > ctx.prev_mfi14_15m:
            mfi_momentum += 0.15
        if ctx.mfi5_15m > ctx.prev_mfi5_15m:
            mfi_momentum += 0.10
        if ctx.mfi14_30m >= ctx.prev_mfi14_30m:
            mfi_momentum += 0.10
        if ctx.mfi14_15m >= 70.0 or ctx.mfi5_15m >= 80.0:
            mfi_momentum -= 0.35  # Overbought penalty

        adjusted_win_rate = min(0.90, max(0.20, self.base_win_rate * mfi_momentum))
        loss_rate = 1.0 - adjusted_win_rate
        
        ev = (adjusted_win_rate * dist_to_target) - (loss_rate * dist_to_sl)
        return float(ev)

    def calculate_trailing_sl(self, ctx: MarketContext, position: Position) -> float:
        """Default trailing SL logic maintaining the non-widening global risk invariant."""
        return position.current_sl


# =====================================================================
# 4. CONCRETE STRATEGY IMPLEMENTATIONS (7 CORE STRATEGIES)
# =====================================================================

class InitialDualMFILowerBandBounceStrategy(BaseStrategy):
    """Strategy 1: Mean-Reversion Dual MFI Bounce near Lower Band."""
    
    def __init__(self):
        super().__init__("Initial Dual MFI Lower Band Bounce", base_win_rate=0.68, base_rr=2.2)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        is_mfi5_zero_bounce = (ctx.prev_mfi5_15m == 0.0 and ctx.mfi5_15m > 0.0) or (ctx.mfi5_15m == 0.0)
        is_mfi14_bounce = (ctx.mfi14_15m <= 25.0) and (ctx.mfi14_15m >= ctx.prev_mfi14_15m)
        is_below_mb = (ctx.close < ctx.mb_20)
        is_30m_ok = (ctx.mfi14_30m >= ctx.prev_mfi14_30m)
        
        if is_mfi5_zero_bounce and is_mfi14_bounce and is_below_mb and is_30m_ok:
            sl = ctx.close - 20.0
            target = ctx.close + max(30.0, (ctx.mb_20 - ctx.close) * 1.5)
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, reason="15m MFI(5)=0 bounce + MFI(14)<=25 rising")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m and ctx.close > position.entry_price:
            return ExitSignal(True, ctx.close, "Dual MFI Falling in Profit Exit")
        if ctx.close >= position.target_price:
            return ExitSignal(True, ctx.close, "Target Hit")
        return ExitSignal(False, ctx.close, "")


class PreviousHighBreakoutMomentumStrategy(BaseStrategy):
    """Strategy 2: Breakout Momentum with Pullback Retest & Sub-MB Bounce."""
    
    def __init__(self):
        super().__init__("Previous High Breakout Momentum Entry", base_win_rate=0.72, base_rr=2.8)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if ctx.is_0915_bar:
            return None  # Suppress millisecond opening spike

        is_both_15m_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m) and (ctx.mfi14_15m >= ctx.prev_mfi14_15m)
        is_30m_both_favorable = (ctx.mfi5_30m >= ctx.prev_mfi5_30m) and (ctx.mfi14_30m >= ctx.prev_mfi14_30m + 0.5)
        is_price_breakout = (ctx.high > ctx.prev_high)
        is_below_mb = (ctx.open < ctx.mb_20)
        
        if is_below_mb:
            valid_entry = is_both_15m_rising and (ctx.close >= ctx.open) and is_price_breakout and is_30m_both_favorable
        else:
            valid_entry = is_both_15m_rising and is_price_breakout and is_30m_both_favorable
            
        if valid_entry:
            prev_candle_range = (ctx.prev_high - ctx.prev_low) if ctx.prev_high > ctx.prev_low else 30.0
            dynamic_pullback = max(6.0, min(18.0, 0.30 * prev_candle_range))
            retest_zone_high = ctx.prev_high
            retest_zone_low = ctx.prev_high - dynamic_pullback
            if ctx.low <= retest_zone_high:
                entry_p = max(retest_zone_low, ctx.low)
            else:
                entry_p = ctx.prev_high + 1.0
                
            entry_both_15m_mfi_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m and ctx.mfi14_15m > ctx.prev_mfi14_15m)
            sl = ctx.low - 15.0 if entry_both_15m_mfi_rising else entry_p - 20.0
            target = entry_p + 40.0
            return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason="Breakout Momentum with 15m & 30m Dual MFI rising")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        is_30m_htf_trend_rising = (ctx.mfi5_30m >= ctx.prev_mfi5_30m and ctx.mfi14_30m >= ctx.prev_mfi14_30m)
        is_price_and_mfi5_rising = (ctx.close > position.entry_price and ctx.mfi5_15m > ctx.prev_mfi5_15m)
        is_htf_hold_trend = is_30m_htf_trend_rising or is_price_and_mfi5_rising

        # 1. Overbought / Upper BB Rejection Exit
        if (ctx.mfi14_15m >= 70.0 or ctx.mfi5_15m >= 80.0) and (ctx.high >= ctx.prev_high or position.peak_price >= ctx.prev_high) and (ctx.close < ctx.prev_high or ctx.close <= ctx.open or ctx.high - ctx.close >= 5.0):
            return ExitSignal(True, ctx.close, "Overbought High Rejection Exit")

        # 2. Upper BB reached + dual MFI falling
        if (ctx.high >= ctx.ub_20 or position.peak_price >= ctx.ub_20) and (ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m):
            return ExitSignal(True, ctx.close, "Upper BB Reached & Dual MFI Fall Exit")

        # 3. MFI(14) or both MFI falling (override if HTF riding)
        if (ctx.mfi14_15m < ctx.prev_mfi14_15m or (ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m)) and not is_htf_hold_trend:
            return ExitSignal(True, ctx.close, "Breakout MFI(14) / Both MFI Fall Exit")
            
        # 4. Weak closing retrace to open (exempt if MFI 14 still rising)
        if position.peak_price >= position.entry_price + 5.0 and (ctx.close <= ctx.open or ctx.high - ctx.close >= 6.0):
            if not (ctx.mfi14_15m > ctx.prev_mfi14_15m):
                return ExitSignal(True, ctx.close, "Breakout Weak Close / Retrace to Open Exit")
                
        # 5. MB Rejection
        if position.peak_price >= ctx.mb_20 - 5.0 and (ctx.mfi14_15m < ctx.prev_mfi14_15m or ctx.mfi5_15m >= 80.0) and ctx.close <= position.peak_price - 3.0:
            return ExitSignal(True, ctx.close, "Breakout Middle Band Rejection Exit")
            
        return ExitSignal(False, ctx.close, "")


class DualMFI30mSecondaryReversalStrategy(BaseStrategy):
    """Strategy 3: 30-min Dual MFI Reversal Option."""
    
    def __init__(self):
        super().__init__("30m Dual MFI Secondary Reversal Option", base_win_rate=0.64, base_rr=2.0)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if not ctx.initial_entry_done:
            return None
        is_15m_both_inc = (ctx.mfi5_15m > ctx.prev_mfi5_15m) and (ctx.mfi14_15m > ctx.prev_mfi14_15m)
        if is_15m_both_inc and ctx.mfi14_15m <= 35.0 and ctx.close >= ctx.open:
            sl = ctx.close - 20.0
            target = ctx.close + 35.0
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, reason="30m MFI Secondary Reversal (MFI14<=35 & Rising)")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_30m < ctx.prev_mfi5_30m and ctx.mfi14_30m < ctx.prev_mfi14_30m and ctx.close > position.entry_price:
            return ExitSignal(True, ctx.close, "30m Dual MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


class MFITrendReentryMBStrategy(BaseStrategy):
    """Strategy 4: MFI(14) Trend Re-Entry (MB Consolidation)."""
    
    def __init__(self):
        super().__init__("MFI(14) Trend Re-Entry (MB Consolidation)", base_win_rate=0.70, base_rr=2.4)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if not ctx.allow_reentry:
            return None
        # Strict overbought & HTF falling blockers
        is_15m_ob = (ctx.mfi5_15m >= 80.0 or ctx.mfi14_15m >= 68.0)
        is_30m_mfi14_falling = (ctx.mfi14_30m < ctx.prev_mfi14_30m) or (ctx.prev_mfi14_30m < ctx.prev_prev_mfi14_30m)
        if is_15m_ob or is_30m_mfi14_falling or ctx.mfi5_30m == 100.0 or ctx.mfi14_30m >= 70.0:
            return None
        
        is_near_mb = (ctx.low <= ctx.mb_20 + 5.0 or ctx.open - ctx.low >= 10.0) and (ctx.close >= ctx.open or ctx.close >= ctx.mb_20)
        is_dual_mfi_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m) and (ctx.mfi14_15m > ctx.prev_mfi14_15m)
        
        if is_near_mb and is_dual_mfi_rising:
            sl = max(ctx.low - 5.0, ctx.close - 20.0)
            target = ctx.close + 35.0
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, reason="MB Consolidation with Dual 15m MFI Rising")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        is_near_ub = (position.peak_price >= ctx.ub_20 - 5.0 or ctx.high >= ctx.ub_20 - 5.0)
        if is_near_ub and (ctx.mfi14_15m < ctx.prev_mfi14_15m or (ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m)):
            return ExitSignal(True, ctx.close, "Re-Entry Upper BB MFI(14) / Dual Fall Exit")
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m:
            return ExitSignal(True, ctx.close, "Re-Entry Dual 15m MFI Fall Strict Exit")
        return ExitSignal(False, ctx.close, "")


class PostSLRecoveryReentryStrategy(BaseStrategy):
    """Strategy 5: Post-SL Recovery Re-Entry (+2 Lots)."""
    
    def __init__(self):
        super().__init__("One-Time Post-SL Recovery Re-Entry (+2 Lots)", base_win_rate=0.74, base_rr=2.5)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        if not ctx.recovery_eligible:
            return None
        is_dual_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m) and (ctx.mfi14_15m > ctx.prev_mfi14_15m)
        is_30m_ok = (ctx.mfi5_30m >= ctx.prev_mfi5_30m) and (ctx.mfi14_30m >= ctx.prev_mfi14_30m + 0.5)
        
        if (ctx.close >= ctx.open) and is_dual_rising and (ctx.close < ctx.mb_20) and is_30m_ok:
            sl = ctx.close - 20.0
            target = ctx.close + 40.0
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, lot_size=2, reason="Post-SL Recovery Bounce")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m:
            return ExitSignal(True, ctx.close, "Recovery Re-Entry 15m Dual MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


class PostBreakdownOversoldBounceStrategy(BaseStrategy):
    """Strategy 6: Post-Breakdown Oversold Bounce."""
    
    def __init__(self):
        super().__init__("Post-Breakdown Oversold Bounce Entry", base_win_rate=0.62, base_rr=2.1)

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        is_prev_breakdown = (ctx.prev_close < ctx.prev_low)
        is_oversold = (ctx.mfi5_15m <= 25.0 or ctx.mfi14_15m <= 30.0)
        is_mfi_rising = (ctx.mfi5_15m > ctx.prev_mfi5_15m or ctx.mfi14_15m > ctx.prev_mfi14_15m)
        
        if is_prev_breakdown and is_oversold and is_mfi_rising:
            sl = ctx.close - 20.0
            target = ctx.close + 30.0
            return EntrySignal(self.name, PositionSide.LONG, ctx.close, sl, target, reason="Oversold Bounce post Breakdown")
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        if (ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m) or (ctx.mfi14_15m < ctx.prev_mfi14_15m and ctx.close < ctx.open):
            return ExitSignal(True, ctx.close, "Post-Breakdown Price Rejection / MFI Fall Exit")
        return ExitSignal(False, ctx.close, "")


class DynamicSwingLowBreakoutRetestStrategy(BaseStrategy):
    """Strategy 3: Dynamic Swing Low Breakout & Retest with MFI confirmation and Upper BB Rejection Exit."""
    
    def __init__(self):
        super().__init__("Dynamic Swing Low First Breakout Retest Entry", base_win_rate=0.74, base_rr=2.5)
        self.waiting_for_bb_pullback: bool = False
        self.lower_bb_touched: bool = False
        self.current_day: str = ""

    def _reset_day_state(self, current_day: str):
        if self.current_day != current_day:
            self.current_day = current_day
            self.waiting_for_bb_pullback = False
            self.lower_bb_touched = False

    def evaluate_entry(self, ctx: MarketContext) -> Optional[EntrySignal]:
        c_day = ctx.timestamp[:10] if len(ctx.timestamp) >= 10 else ""
        self._reset_day_state(c_day)

        is_mfi_falling_from_ob = (ctx.prev_mfi5_15m >= 80.0 or ctx.prev_mfi14_15m >= 70.0) and (ctx.mfi5_15m < ctx.prev_mfi5_15m or ctx.mfi14_15m < ctx.prev_mfi14_15m)
        is_ob_blocked = (ctx.mfi5_15m >= 80.0 or ctx.mfi14_15m >= 68.0) or (ctx.ub_20 > 0 and ctx.high >= ctx.ub_20 - 8.0) or is_mfi_falling_from_ob
        if is_ob_blocked:
            self.waiting_for_bb_pullback = False
            self.lower_bb_touched = False
            return None

        swing_low_bounce_zone = max(25.0, ctx.dynamic_tolerance * 2.5)
        
        # Middle Band Proximity & Correction at or below MB:
        mb_level = ctx.mb_3m if ctx.mb_3m > 0 else ctx.mb_20
        is_corrected_to_mb = (mb_level > 0) and (ctx.low <= mb_level + 1.0)
        is_bounce_to_open = (ctx.close >= ctx.open)
        is_mb_setup = is_corrected_to_mb and is_bounce_to_open

        is_near_swing_low = (ctx.low <= ctx.recent_swing_low + swing_low_bounce_zone) or (ctx.open <= ctx.recent_swing_low + swing_low_bounce_zone)
        is_candidate_setup = (is_mb_setup or is_near_swing_low) and is_bounce_to_open

        # 3min MFI indicators
        curr_mfi5_3m = ctx.mfi5_3m if ctx.mfi5_3m > 0 else ctx.mfi5_15m
        prev_mfi5_3m = ctx.prev_mfi5_3m if ctx.prev_mfi5_3m > 0 else ctx.prev_mfi5_15m
        prev_prev_mfi5_3m = ctx.prev_prev_mfi5_3m if ctx.prev_prev_mfi5_3m > 0 else prev_mfi5_3m
        curr_mfi14_3m = ctx.mfi14_3m if ctx.mfi14_3m > 0 else ctx.mfi14_15m
        prev_mfi14_3m = ctx.prev_mfi14_3m if ctx.prev_mfi14_3m > 0 else ctx.prev_mfi14_15m
        prev_prev_mfi14_3m = ctx.prev_prev_mfi14_3m if ctx.prev_prev_mfi14_3m > 0 else prev_mfi14_3m

        # Rule: During entry time or in previous candle if 3min both MFI falling
        curr_3m_both_falling = (curr_mfi5_3m < prev_mfi5_3m) and (curr_mfi14_3m < prev_mfi14_3m)
        prev_3m_both_falling = (prev_mfi5_3m < prev_prev_mfi5_3m) and (prev_mfi14_3m < prev_prev_mfi14_3m)
        is_3m_both_falling = curr_3m_both_falling or prev_3m_both_falling

        # Required Reversal Criteria: Price touches lower BB, both 3m MFI rise, HTF MFI rise
        lb_target = ctx.lb_3m if ctx.lb_3m > 0 else ctx.lb_20
        if lb_target > 0 and ctx.low <= lb_target + 2.0:
            self.lower_bb_touched = True

        # 3m MFI Rising check: Both MFI increasing OR at least MFI(14) increasing
        is_3m_mfi_rising = ((curr_mfi5_3m > prev_mfi5_3m) and (curr_mfi14_3m >= prev_mfi14_3m)) or (curr_mfi14_3m > prev_mfi14_3m)
        is_15m_mfi_rising = (ctx.mfi14_15m > ctx.prev_mfi14_15m) or (ctx.mfi14_15m >= ctx.prev_mfi14_15m and ctx.mfi5_15m > ctx.prev_mfi5_15m)
        is_30m_mfi_rising = (ctx.mfi14_30m > ctx.prev_mfi14_30m) or (ctx.mfi14_30m >= ctx.prev_mfi14_30m and ctx.mfi5_30m > ctx.prev_mfi5_30m)
        is_htf_mfi_rising = is_15m_mfi_rising or is_30m_mfi_rising

        # Handle 3min falling wait state: wait for price to touch lower BB and 3m MFI to rise and HTF MFI rise
        if self.waiting_for_bb_pullback:
            if self.lower_bb_touched and is_3m_mfi_rising and is_htf_mfi_rising and is_bounce_to_open:
                self.waiting_for_bb_pullback = False
                self.lower_bb_touched = False
                retest_trigger = ctx.recent_swing_low + swing_low_bounce_zone
                entry_p = ctx.open if ctx.open <= retest_trigger else min(ctx.open, retest_trigger)
                sl = max(ctx.low - 5.0, entry_p - 20.0)
                target = entry_p + 35.0
                return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason="Dynamic Swing Low / Lower BB Touch & 3m+HTF MFI Reversal")
            return None

        if is_candidate_setup:
            if is_3m_both_falling:
                # 3m both MFI falling at entry time or prev candle -> activate wait for lower BB touch + 3m & HTF MFI rise
                self.waiting_for_bb_pullback = True
                if self.lower_bb_touched and is_3m_mfi_rising and is_htf_mfi_rising and is_bounce_to_open:
                    self.waiting_for_bb_pullback = False
                    self.lower_bb_touched = False
                    retest_trigger = ctx.recent_swing_low + swing_low_bounce_zone
                    entry_p = ctx.open if ctx.open <= retest_trigger else min(ctx.open, retest_trigger)
                    sl = max(ctx.low - 5.0, entry_p - 20.0)
                    target = entry_p + 35.0
                    return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason="Dynamic Swing Low / Lower BB Touch & 3m+HTF MFI Reversal")
                return None
            else:
                # Can enter from middle band only if price corrected at/below MB, bounced to Open, both 3m MFI (or MFI14) rising AND any HTF MFI rising
                if is_3m_mfi_rising and is_htf_mfi_rising and is_bounce_to_open:
                    retest_trigger = ctx.recent_swing_low + swing_low_bounce_zone
                    entry_p = ctx.open if ctx.open <= retest_trigger else min(ctx.open, retest_trigger)
                    sl = max(ctx.low - 5.0, entry_p - 20.0)
                    target = entry_p + 35.0
                    reason_str = "Dynamic Swing Low / Middle Band Entry (3m+HTF MFI Rising)" if is_mb_setup else "Dynamic Swing Low / Retest Bounce (3m+HTF MFI Rising)"
                    return EntrySignal(self.name, PositionSide.LONG, entry_p, sl, target, reason=reason_str)
        return None

    def evaluate_exit(self, ctx: MarketContext, position: Position) -> ExitSignal:
        # 3m Upper BB / MFI Profit Booking Exit Logic (Exact User Specification)
        curr_mfi5_3m = ctx.mfi5_3m if ctx.mfi5_3m > 0 else ctx.mfi5_15m
        prev_mfi5_3m = ctx.prev_mfi5_3m if ctx.prev_mfi5_3m > 0 else ctx.prev_mfi5_15m
        curr_mfi14_3m = ctx.mfi14_3m if ctx.mfi14_3m > 0 else ctx.mfi14_15m
        prev_mfi14_3m = ctx.prev_mfi14_3m if ctx.prev_mfi14_3m > 0 else ctx.prev_mfi14_15m

        # 1. Hold position if both 3m MFIs are increasing by >1.0 point
        both_3m_mfi_surging = (curr_mfi5_3m >= prev_mfi5_3m + 1.0) and (curr_mfi14_3m >= prev_mfi14_3m + 1.0)
        # OR if MFI(5) reaches 100 while MFI(14) is still rising
        mfi100_and_14_rising = (curr_mfi5_3m >= 99.0 or ctx.mfi5_15m >= 99.0) and (curr_mfi14_3m >= prev_mfi14_3m or ctx.mfi14_15m >= ctx.prev_mfi14_15m)
        hold_due_to_surging_mfi = both_3m_mfi_surging or mfi100_and_14_rising

        # 2. Exit on 3m Upper BB touch / +40pt gain ONLY IF 3m MFI is falling AND HTF MFI is falling
        is_mfi14_falling_3m = (curr_mfi14_3m < prev_mfi14_3m)
        both_mfi_falling_3m = (curr_mfi5_3m < prev_mfi5_3m and curr_mfi14_3m < prev_mfi14_3m)
        exit_mfi_confirmed = is_mfi14_falling_3m or both_mfi_falling_3m

        # HTF MFI Hierarchy check: Only exit if higher timeframe MFI (15m/30m) is also falling
        is_htf_mfi_falling = (ctx.mfi14_15m < ctx.prev_mfi14_15m) or (ctx.mfi14_30m < ctx.prev_mfi14_30m)

        ub_target = ctx.ub_3m if ctx.ub_3m > 0 else ctx.ub_20
        is_3m_ub_near = (ub_target > 0) and (ctx.high >= ub_target - 1.5 or ctx.close >= ub_target - 1.5 or position.peak_price >= ub_target - 1.5)
        is_40pt_gain = (position.peak_price >= position.entry_price + 40.0) or (ctx.high >= position.entry_price + 40.0) or (ctx.close >= position.entry_price + 40.0)

        if (is_3m_ub_near or is_40pt_gain) and exit_mfi_confirmed and is_htf_mfi_falling and not hold_due_to_surging_mfi:
            if ub_target > 0 and ctx.high >= ub_target:
                exit_price = max(ctx.open, min(ctx.high, ub_target))
            elif ctx.high >= position.entry_price + 40.0:
                exit_price = max(ctx.open, position.entry_price + 40.0)
            else:
                exit_price = ctx.close
            return ExitSignal(True, exit_price, "3m Upper BB / +40pt Profit Booking (HTF MFI Falling Confirmed)")

        if hold_due_to_surging_mfi:
            return ExitSignal(False, ctx.close, "")  # Waiting / Holding position while 3m MFI is surging or MFI(5)=100 & MFI(14) rising

        # Standard strategy fallback exits when MFI is not actively surging:
        if (position.peak_price >= ctx.ub_20 - 5.0 or ctx.high >= ctx.ub_20 - 5.0) and (ctx.mfi14_15m >= 70.0 or ctx.mfi5_15m >= 80.0) and (ctx.close <= ctx.open):
            return ExitSignal(True, ctx.close, "Swing Low Overbought Retrace Exit")
        if ctx.mfi5_15m < ctx.prev_mfi5_15m and ctx.mfi14_15m < ctx.prev_mfi14_15m:
            return ExitSignal(True, ctx.close, "Swing Low Dual MFI Fall Exit")
        if ctx.close >= position.target_price:
            return ExitSignal(True, ctx.close, "Target Hit (RR Booked)")
        return ExitSignal(False, ctx.close, "")

    def calculate_trailing_sl(self, ctx: MarketContext, position: Position) -> float:
        """Dynamic trailing SL: multi-stage profit lock starting above 40 points.
        No trail before 40 points as long as MFI(14) rising in 15 minutes.
        """
        favorable = position.peak_price - position.entry_price if position.side == PositionSide.LONG else position.entry_price - position.peak_price
        sl = position.current_sl
        is_mfi14_rising = (ctx.mfi14_15m > ctx.prev_mfi14_15m) or (ctx.mfi14_15m >= ctx.prev_mfi14_15m and ctx.mfi5_15m > ctx.prev_mfi5_15m)

        if is_mfi14_rising:
            # Rule: No trail before 40 points while MFI(14) is rising in 15m. Trail starts above 40 points:
            if favorable >= 48.0:
                sl = max(sl, position.entry_price + 18.0 if position.side == PositionSide.LONG else position.entry_price - 18.0)
            elif favorable >= 40.0:
                sl = max(sl, position.entry_price + 10.0 if position.side == PositionSide.LONG else position.entry_price - 10.0)
        else:
            if favorable >= 40.0:
                sl = max(sl, position.entry_price + 10.0 if position.side == PositionSide.LONG else position.entry_price - 10.0)
            elif favorable >= 18.0:
                sl = max(sl, position.entry_price + 10.0 if position.side == PositionSide.LONG else position.entry_price - 10.0)
            elif favorable >= 10.0:
                sl = max(sl, position.entry_price + 3.0 if position.side == PositionSide.LONG else position.entry_price - 3.0)
        return sl

    def calculate_expected_value(self, ctx: MarketContext, position: Position) -> float:
        """Dynamic EV calculation with MFI(14) breakout velocity weighting."""
        mfi_momentum = 1.0
        if ctx.mfi14_15m > ctx.prev_mfi14_15m:
            mfi_momentum += 0.25
        if ctx.mfi5_15m > ctx.prev_mfi5_15m:
            mfi_momentum += 0.15
        if ctx.mfi14_15m >= 70.0 and ctx.high >= ctx.ub_20 - 2.0:
            mfi_momentum -= 0.30

        dist_to_target = max(5.0, position.target_price - ctx.close)
        dist_to_sl = max(5.0, ctx.close - position.current_sl)
        adjusted_win_rate = min(0.92, max(0.20, self.base_win_rate * mfi_momentum))
        loss_rate = 1.0 - adjusted_win_rate
        return float((adjusted_win_rate * dist_to_target) - (loss_rate * dist_to_sl))


# =====================================================================
# 5. STATE MACHINE & IN-FLIGHT HANDOVER ENGINE
# =====================================================================

class StateMachineHandoverEngine:
    """
    Core State Machine that manages active trades and evaluates mid-flight
    handover opportunities between strategies based on dynamic regime & EV.
    """

    def __init__(self, strategies: List[BaseStrategy], min_ev_improvement: float = 4.0):
        self.strategies: Dict[str, BaseStrategy] = {s.name: s for s in strategies}
        self.min_ev_improvement = min_ev_improvement
        self.state = MachineState.STATE_0_SCANNING
        self.active_position: Optional[Position] = None
        self.trade_counter = 0
        self.closed_trades: List[Dict[str, Any]] = []

    def process_candle(self, ctx: MarketContext) -> Optional[Dict[str, Any]]:
        """Processes a single candle through the state machine."""
        time_str = ctx.timestamp.split("T")[-1].split(" ")[-1][:5] if ("T" in ctx.timestamp or " " in ctx.timestamp) else ""

        # UNIVERSAL RULE #1: Hard Blocker on MFI(5)=100 or extreme overbought
        is_universal_ob = (ctx.mfi5_15m >= 99.0) or (ctx.mfi14_15m >= 70.0) or (ctx.mfi5_15m >= 90.0 and not (ctx.mfi14_15m > ctx.prev_mfi14_15m))
        
        # -------------------------------------------------------------
        # STATE 0: SCANNING / FLAT
        # -------------------------------------------------------------
        if self.state == MachineState.STATE_0_SCANNING:
            # Rule: Restrict fresh Entry on or after 15:00:00
            if time_str >= "15:00":
                return None

            if is_universal_ob:
                return None  # Block fresh entries at 100/overbought

            # Synchronized scan across all candidate strategies
            for strategy in self.strategies.values():
                signal = strategy.evaluate_entry(ctx)
                if signal:
                    self.trade_counter += 1
                    self.active_position = Position(
                        trade_id=self.trade_counter,
                        strategy_name=signal.strategy_name,
                        side=signal.side,
                        entry_time=ctx.timestamp,
                        entry_price=signal.entry_price,
                        current_sl=signal.initial_sl,
                        initial_sl=signal.initial_sl,
                        target_price=signal.target_price,
                        peak_price=signal.entry_price,
                        lot_size=signal.lot_size,
                        metadata=signal.metadata
                    )
                    self.state = MachineState.STATE_1_ACTIVE_LONG if signal.side == PositionSide.LONG else MachineState.STATE_2_ACTIVE_SHORT
                    return {
                        "event": "ENTRY",
                        "trade_id": self.active_position.trade_id,
                        "strategy": self.active_position.strategy_name,
                        "time": ctx.timestamp,
                        "price": self.active_position.entry_price,
                        "sl": self.active_position.current_sl,
                        "target": self.active_position.target_price,
                        "reason": signal.reason
                    }
            return None

        # -------------------------------------------------------------
        # STATE 1 & 2: ACTIVE POSITION (EVALUATE SL, EXITS & HANDOVER)
        # -------------------------------------------------------------
        if self.state in (MachineState.STATE_1_ACTIVE_LONG, MachineState.STATE_2_ACTIVE_SHORT, MachineState.STATE_3_HANDOVER_WINDOW):
            pos = self.active_position
            if not pos:
                self.state = MachineState.STATE_0_SCANNING
                return None

            # Update peak price
            if pos.side == PositionSide.LONG:
                pos.peak_price = max(pos.peak_price, ctx.high)
            else:
                pos.peak_price = min(pos.peak_price, ctx.low)

            current_strat = self.strategies.get(pos.strategy_name)
            if not current_strat:
                current_strat = list(self.strategies.values())[0]

            # Check Hard Stop Loss Hit
            if (pos.side == PositionSide.LONG and ctx.low <= pos.current_sl) or (pos.side == PositionSide.SHORT and ctx.high >= pos.current_sl):
                exit_record = self._close_position(ctx, pos, exit_price=pos.current_sl, reason="Stop Loss Hit")
                return exit_record

            # Trailing SL: Move SL to Cost - 7pts on +20pts favorable move
            favorable_move = pos.peak_price - pos.entry_price if pos.side == PositionSide.LONG else pos.entry_price - pos.peak_price
            if favorable_move >= 20.0:
                pos.current_sl = max(pos.current_sl, pos.entry_price - 7.0)

            # Mandatory EOD Cutoff Exit at 15:25
            time_str = ctx.timestamp.split("T")[-1].split(" ")[-1][:5] if ("T" in ctx.timestamp or " " in ctx.timestamp) else ""
            if time_str >= "15:25":
                exit_record = self._close_position(ctx, pos, exit_price=ctx.close, reason="EOD 15:25 Cutoff Exit")
                return exit_record

            # Dynamic Trailing SL from Current Strategy
            pos.current_sl = max(pos.current_sl, current_strat.calculate_trailing_sl(ctx, pos))

            # ---------------------------------------------------------
            # STATE 3: IN-FLIGHT DYNAMIC STRATEGY HANDOVER EVALUATION
            # ---------------------------------------------------------
            handover = self._evaluate_in_flight_handover(ctx, pos)
            if handover.switch_approved:
                self.state = MachineState.STATE_3_HANDOVER_WINDOW
                pos.handover_history.append({
                    "time": ctx.timestamp,
                    "from_strategy": handover.from_strategy,
                    "to_strategy": handover.to_strategy,
                    "reason": handover.reason,
                    "projected_ev": handover.projected_ev
                })
                pos.strategy_name = handover.to_strategy
                pos.current_sl = max(pos.current_sl, handover.adjusted_sl)  # Global Risk Gate: Never widen SL
                pos.target_price = handover.adjusted_target
                current_strat = self.strategies[handover.to_strategy]

            # Evaluate Active Strategy Exit Condition
            exit_sig = current_strat.evaluate_exit(ctx, pos)
            if exit_sig.should_exit:
                exit_record = self._close_position(ctx, pos, exit_price=exit_sig.exit_price, reason=exit_sig.exit_reason)
                return exit_record

        return None

    def _evaluate_in_flight_handover(self, ctx: MarketContext, pos: Position) -> HandoverDecision:
        """
        Calculates EV across all candidate strategies and executes dynamic
        handover if a regime change (e.g. Mean Reversion -> Strong Trend) justifies switching.
        """
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
            # Check Regime Transition: e.g. Mean-Reversion -> Trend Breakout riding
            new_target = pos.target_price
            new_sl = pos.current_sl

            # Extend target and tighten trailing SL for Trend Breakout Handover
            if "Breakout" in best_candidate.name and ctx.mfi14_15m > ctx.prev_mfi14_15m:
                new_target = max(pos.target_price, ctx.close + 50.0)
                new_sl = max(pos.current_sl, ctx.low - 10.0)

            # GLOBAL RISK GATE: Never widen initial risk budget
            guaranteed_sl = max(pos.initial_sl, new_sl)

            return HandoverDecision(
                switch_approved=True,
                from_strategy=pos.strategy_name,
                to_strategy=best_candidate.name,
                current_ev=current_ev,
                projected_ev=best_ev,
                adjusted_sl=guaranteed_sl,
                adjusted_target=new_target,
                reason=f"Dynamic In-Flight Handover: EV improved from {current_ev:.1f} to {best_ev:.1f}pts via {best_candidate.name}"
            )

        return HandoverDecision(False, pos.strategy_name, pos.strategy_name, current_ev, current_ev, pos.current_sl, pos.target_price, "")

    def _close_position(self, ctx: MarketContext, pos: Position, exit_price: float, reason: str) -> Dict[str, Any]:
        """Closes position, logs performance metrics, and resets state."""
        pnl = (exit_price - pos.entry_price) * pos.lot_size if pos.side == PositionSide.LONG else (pos.entry_price - exit_price) * pos.lot_size
        pnl_rounded = round(pnl, 2)
        res_str = "PROFIT" if pnl_rounded > 0 else ("BREAKEVEN" if pnl_rounded == 0 else "LOSS")
        record = {
            "event": "EXIT",
            "trade_id": pos.trade_id,
            "contract": pos.metadata.get("contract", ""),
            "strategy": pos.strategy_name,
            "type": pos.strategy_name,
            "side": pos.side.value,
            "entry_time": pos.entry_time,
            "exit_time": ctx.timestamp,
            "entry": round(pos.entry_price, 2),
            "entry_price": round(pos.entry_price, 2),
            "exit": round(exit_price, 2),
            "exit_price": round(exit_price, 2),
            "pnl": pnl_rounded,
            "result": res_str,
            "handovers": len(pos.handover_history),
            "reason": reason,
            "exit_type": reason,
            "handover_history": pos.handover_history
        }
        self.closed_trades.append(record)
        self.active_position = None
        self.state = MachineState.STATE_0_SCANNING
        return record


# =====================================================================
# 6. UNIFIED HISTORICAL BACKTEST ENGINE & PERFORMANCE REPORTING
# =====================================================================

def run_state_machine_backtest(df_15m: pd.DataFrame, df_30m: Optional[pd.DataFrame] = None, df_3m: Optional[pd.DataFrame] = None, from_date_str: Optional[str] = None, contract_symbol: str = "") -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Executes concurrent synchronized backtest of the 7-strategy suite with Dynamic Handover.
    """
    if df_15m is None or df_15m.empty:
        return pd.DataFrame(), {"total_trades": 0, "win_rate_pct": 0.0, "total_pnl_pts": 0.0, "total_in_flight_handovers": 0, "profit_factor": 0.0}

    # 1. Feature Engineering on Full Warmup Dataset
    df_15m_calc = df_15m.copy()
    for col in ['open', 'high', 'low', 'close', 'volume']:
        if col in df_15m_calc.columns:
            df_15m_calc[col] = pd.to_numeric(df_15m_calc[col], errors='coerce').fillna(0.0)

    # Filter out zero-volume or flat un-traded candles before indicator calculations
    df_15m_calc = df_15m_calc[(df_15m_calc['volume'] > 0.0) & ((df_15m_calc['high'] - df_15m_calc['low']) >= 0.5)].copy().reset_index(drop=True)

    df_15m_calc['mfi5'] = calculate_mfi(df_15m_calc['high'], df_15m_calc['low'], df_15m_calc['close'], df_15m_calc['volume'], period=5)
    df_15m_calc['mfi14'] = calculate_mfi(df_15m_calc['high'], df_15m_calc['low'], df_15m_calc['close'], df_15m_calc['volume'], period=14)
    df_15m_calc['prev_mfi5'] = df_15m_calc['mfi5'].shift(1).fillna(50.0)
    df_15m_calc['prev_mfi14'] = df_15m_calc['mfi14'].shift(1).fillna(50.0)
    
    mb, ub, lb = calculate_bollinger_bands(df_15m_calc['close'], period=20, num_std=2.0)
    df_15m_calc['mb_20'] = mb
    df_15m_calc['ub_20'] = ub
    df_15m_calc['lb_20'] = lb
    df_15m_calc['prev_high'] = df_15m_calc['high'].shift(1).fillna(df_15m_calc['high'])
    df_15m_calc['prev_low'] = df_15m_calc['low'].shift(1).fillna(df_15m_calc['low'])
    df_15m_calc['prev_close'] = df_15m_calc['close'].shift(1).fillna(df_15m_calc['close'])

    # 30m alignment
    if df_30m is not None and not df_30m.empty:
        df_30m_calc = df_30m.copy()
        for col in ['open', 'high', 'low', 'close', 'volume']:
            if col in df_30m_calc.columns:
                df_30m_calc[col] = pd.to_numeric(df_30m_calc[col], errors='coerce').fillna(0.0)
        df_30m_calc['mfi5_30m'] = calculate_mfi(df_30m_calc['high'], df_30m_calc['low'], df_30m_calc['close'], df_30m_calc['volume'], period=5)
        df_30m_calc['mfi14_30m'] = calculate_mfi(df_30m_calc['high'], df_30m_calc['low'], df_30m_calc['close'], df_30m_calc['volume'], period=14)
        df_30m_calc['prev_mfi5_30m'] = df_30m_calc['mfi5_30m'].shift(1).fillna(50.0)
        df_30m_calc['prev_mfi14_30m'] = df_30m_calc['mfi14_30m'].shift(1).fillna(50.0)
        df_30m_calc['prev_prev_mfi14_30m'] = df_30m_calc['mfi14_30m'].shift(2).fillna(50.0)
    else:
        df_30m_calc = df_15m_calc.copy()
        df_30m_calc['mfi5_30m'] = df_30m_calc['mfi5']
        df_30m_calc['mfi14_30m'] = df_30m_calc['mfi14']
        df_30m_calc['prev_mfi5_30m'] = df_30m_calc['prev_mfi5']
        df_30m_calc['prev_mfi14_30m'] = df_30m_calc['prev_mfi14']
        df_30m_calc['prev_prev_mfi14_30m'] = df_30m_calc['mfi14_30m'].shift(1).fillna(50.0)

    # 3m alignment for micro-timeframe 3m bar resolution execution
    if df_3m is not None and not df_3m.empty:
        df_3m_calc = df_3m.copy()
        for col in ['open', 'high', 'low', 'close', 'volume']:
            if col in df_3m_calc.columns:
                df_3m_calc[col] = pd.to_numeric(df_3m_calc[col], errors='coerce').fillna(0.0)
        df_3m_calc = df_3m_calc[(df_3m_calc['volume'] > 0.0) & ((df_3m_calc['high'] - df_3m_calc['low']) >= 0.5)].copy().reset_index(drop=True)
        df_3m_calc['mfi5_3m'] = calculate_mfi(df_3m_calc['high'], df_3m_calc['low'], df_3m_calc['close'], df_3m_calc['volume'], period=5)
        df_3m_calc['mfi14_3m'] = calculate_mfi(df_3m_calc['high'], df_3m_calc['low'], df_3m_calc['close'], df_3m_calc['volume'], period=14)
        mb_3m, ub_3m, lb_3m = calculate_bollinger_bands(df_3m_calc['close'], period=20, num_std=2.0)
        df_3m_calc['mb_3m'] = mb_3m
        df_3m_calc['ub_3m'] = ub_3m
        df_3m_calc['lb_3m'] = lb_3m
        df_3m_calc['prev_mfi5_3m'] = df_3m_calc['mfi5_3m'].shift(1).fillna(50.0)
        df_3m_calc['prev_mfi14_3m'] = df_3m_calc['mfi14_3m'].shift(1).fillna(50.0)
        df_3m_calc['prev_prev_mfi5_3m'] = df_3m_calc['mfi5_3m'].shift(2).fillna(50.0)
        df_3m_calc['prev_prev_mfi14_3m'] = df_3m_calc['mfi14_3m'].shift(2).fillna(50.0)

        df_eval = pd.merge_ordered(df_3m_calc, df_15m_calc[['timestamp', 'mfi5', 'mfi14', 'prev_mfi5', 'prev_mfi14', 'mb_20', 'ub_20', 'lb_20', 'prev_high', 'prev_low', 'prev_close']], on='timestamp', how='left').ffill()
        df_eval = pd.merge_ordered(df_eval, df_30m_calc[['timestamp', 'mfi5_30m', 'mfi14_30m', 'prev_mfi5_30m', 'prev_mfi14_30m', 'prev_prev_mfi14_30m']], on='timestamp', how='left').ffill()
        df = df_eval.dropna(subset=['open', 'close']).reset_index(drop=True)
    else:
        df = df_15m_calc.copy()
        df['mfi5_3m'] = df['mfi5']
        df['mfi14_3m'] = df['mfi14']
        df['mb_3m'] = df['mb_20']
        df['ub_3m'] = df['ub_20']
        df['lb_3m'] = df['lb_20']
        df['prev_mfi5_3m'] = df['prev_mfi5']
        df['prev_mfi14_3m'] = df['prev_mfi14']
        df['prev_prev_mfi5_3m'] = df['prev_mfi5']
        df['prev_prev_mfi14_3m'] = df['prev_mfi14']
        df = pd.merge_ordered(df, df_30m_calc[['timestamp', 'mfi5_30m', 'mfi14_30m', 'prev_mfi5_30m', 'prev_mfi14_30m', 'prev_prev_mfi14_30m']], on='timestamp', how='left').ffill()

    df['mfi5_3m'] = df['mfi5_3m'].fillna(df['mfi5'])
    df['mfi14_3m'] = df['mfi14_3m'].fillna(df['mfi14'])
    df['mb_3m'] = df['mb_3m'].fillna(df['mb_20'])
    df['ub_3m'] = df['ub_3m'].fillna(df['ub_20'])
    df['lb_3m'] = df['lb_3m'].fillna(df['lb_20'])
    df['prev_mfi5_3m'] = df['prev_mfi5_3m'].fillna(df['prev_mfi5'])
    df['prev_mfi14_3m'] = df['prev_mfi14_3m'].fillna(df['prev_mfi14'])
    df['prev_prev_mfi5_3m'] = df['prev_prev_mfi5_3m'].fillna(df['prev_mfi5'])
    df['prev_prev_mfi14_3m'] = df['prev_prev_mfi14_3m'].fillna(df['prev_mfi14'])

    # Instantiate the Modular 3-Strategy Suite
    strategies = [
        PreviousHighBreakoutMomentumStrategy(),
        PostSLRecoveryReentryStrategy(),
        DynamicSwingLowBreakoutRetestStrategy(),
    ]

    engine = StateMachineHandoverEngine(strategies=strategies, min_ev_improvement=3.5)
    trade_events = []
    initial_entry_done = False
    recovery_eligible = False
    sl_stopped_dates = set()

    for i in range(len(df)):
        row = df.iloc[i]
        ts_str = str(row['timestamp'])
        c_date = ts_str[:10]

        # Ignore warmup candles prior to from_date_str for trade triggering
        if from_date_str and c_date < from_date_str:
            continue

        # If Stop Loss was hit for this side today, halt new entries for this day
        if not engine.active_position and c_date in sl_stopped_dates:
            continue

        lookback = min(7, i) if i >= 1 else 1
        recent_low = df['low'].iloc[max(0, i - lookback):i].min() if i >= 1 else row['low']
        recent_high = df['high'].iloc[max(0, i - lookback):i].max() if i >= 1 else row['high']
        dyn_range = max(15.0, recent_high - recent_low)
        dyn_tol = max(3.0, min(12.0, 0.15 * dyn_range))

        is_0915 = "09:15" in ts_str

        ctx = MarketContext(
            timestamp=ts_str,
            open=float(row['open']),
            high=float(row['high']),
            low=float(row['low']),
            close=float(row['close']),
            volume=float(row['volume']),
            mfi5_15m=float(row['mfi5']),
            mfi14_15m=float(row['mfi14']),
            prev_mfi5_15m=float(row['prev_mfi5']),
            prev_mfi14_15m=float(row['prev_mfi14']),
            mfi5_30m=float(row['mfi5_30m']),
            mfi14_30m=float(row['mfi14_30m']),
            prev_mfi5_30m=float(row['prev_mfi5_30m']),
            prev_mfi14_30m=float(row['prev_mfi14_30m']),
            prev_prev_mfi14_30m=float(row['prev_prev_mfi14_30m']),
            mfi5_60m=50.0,
            mfi14_60m=50.0,
            prev_mfi5_60m=50.0,
            prev_mfi14_60m=50.0,
            mfi5_3m=float(row['mfi5_3m']),
            mfi14_3m=float(row['mfi14_3m']),
            prev_mfi5_3m=float(row['prev_mfi5_3m']),
            prev_mfi14_3m=float(row['prev_mfi14_3m']),
            prev_prev_mfi5_3m=float(row['prev_prev_mfi5_3m']),
            prev_prev_mfi14_3m=float(row['prev_prev_mfi14_3m']),
            mb_3m=float(row['mb_3m']),
            ub_3m=float(row['ub_3m']),
            lb_3m=float(row['lb_3m']),
            mb_20=float(row['mb_20']),
            ub_20=float(row['ub_20']),
            lb_20=float(row['lb_20']),
            prev_high=float(row['prev_high']),
            prev_low=float(row['prev_low']),
            prev_close=float(row['prev_close']),
            recent_swing_low=float(recent_low),
            recent_swing_high=float(recent_high),
            dynamic_tolerance=float(dyn_tol),
            is_0915_bar=is_0915,
            is_big_gap_up=False,
            allow_reentry=True,
            recovery_eligible=recovery_eligible,
            initial_entry_done=initial_entry_done
        )

        event = engine.process_candle(ctx)
        if event:
            trade_events.append(event)
            if event["event"] == "ENTRY":
                initial_entry_done = True
                if engine.active_position and contract_symbol:
                    engine.active_position.metadata["contract"] = contract_symbol
            elif event["event"] == "EXIT" and event["pnl"] < 0:
                sl_stopped_dates.add(c_date)
                recovery_eligible = False

    # Performance Analytics
    closed_df = pd.DataFrame(engine.closed_trades)
    if not closed_df.empty:
        total_pnl = closed_df['pnl'].sum()
        wins = (closed_df['pnl'] > 0).sum()
        win_rate = (wins / len(closed_df)) * 100.0
        total_handovers = closed_df['handovers'].sum()
        gross_profit = closed_df[closed_df['pnl'] > 0]['pnl'].sum()
        gross_loss = abs(closed_df[closed_df['pnl'] < 0]['pnl'].sum())
        profit_factor = round(gross_profit / max(1.0, gross_loss), 2) if gross_loss > 0 else (round(gross_profit, 2) if gross_profit > 0 else 1.0)
        summary = {
            "total_trades": len(closed_df),
            "win_rate_pct": round(win_rate, 2),
            "total_pnl_pts": round(total_pnl, 2),
            "total_in_flight_handovers": int(total_handovers),
            "profit_factor": profit_factor
        }
    else:
        summary = {"total_trades": 0, "win_rate_pct": 0.0, "total_pnl_pts": 0.0, "total_in_flight_handovers": 0, "profit_factor": 0.0}

    return closed_df, summary


if __name__ == "__main__":
    print("🚀 [MODULAR 3-STRATEGY SUITE WITH DYNAMIC IN-FLIGHT HANDOVER] Initialized successfully with 3 Core Strategies.")

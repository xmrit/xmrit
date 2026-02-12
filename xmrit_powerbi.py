###############################################################################
# XMRIT X Chart for Power BI - Python Visual
# -------------------------------------------
# Paste this script into a Power BI Python visual.
# Add a Date column and a Value (measure) column to the visual's fields.
#
# Power BI provides data as a pandas DataFrame called `dataset`.
# The first column should be dates, the second should be numeric values.
###############################################################################

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from datetime import datetime, timedelta

# =============================================================================
# 1. CONFIGURATION
# =============================================================================

DECIMAL_PLACES = 2

# Scaling constants (standard XmR)
NPL_SCALING = 2.66
URL_SCALING = 3.268

# Median-based scaling (used when USE_MEDIAN = True)
MEDIAN_NPL_SCALING = 3.145
MEDIAN_URL_SCALING = 3.865
USE_MEDIAN = False              # True = use median for avg and MR; False = use mean

# Feature toggles
SHOW_TREND_LINES = False        # Enable linear regression trend lines
SHOW_QUARTILE_LINES = True      # Show quartile (dashed grey) lines

# Dividers: list of date strings to split data into segments (empty = no dividers)
# Example: ["2025-06-01", "2025-09-01"]
DIVIDER_DATES = []

# Locked Limits: set to None to use calculated values, or provide a number
LOCKED_AVG = None               # e.g. 42.5
LOCKED_UNPL = None              # e.g. 55.0
LOCKED_LNPL = None              # e.g. 30.0

# Seasonality: set SEASONAL_PERIOD to None to disable
# Valid values: "year", "quarter", "month", "week"
SEASONAL_PERIOD = None

# Chart sizing
PADDING_FROM_EXTREMES = 0.1     # 10% padding above/below data extremes

# =============================================================================
# 2. COLORS & STYLE CONSTANTS (matching XMRIT)
# =============================================================================

MEAN_COLOR = "red"
LIMIT_COLOR = "steelblue"
QUARTILE_COLOR = "gray"
DIVIDER_COLOR = "purple"
DATA_LINE_COLOR = "#000000"

# Exception status codes
NORMAL = 0
RUN_OF_EIGHT = 1
NEAR_LIMIT = 2
OUTSIDE_NPL = 3

# Point colors by status
POINT_COLORS = {
    NORMAL: "black",
    RUN_OF_EIGHT: "blue",
    NEAR_LIMIT: "orange",
    OUTSIDE_NPL: "red",
}

# Label colors by status (slightly different for contrast on white bg)
LABEL_COLORS = {
    NORMAL: "black",
    RUN_OF_EIGHT: "blue",
    NEAR_LIMIT: "#be5504",     # ginger
    OUTSIDE_NPL: "#e3242b",    # rose
}

LIMIT_LINE_WIDTH = 2
LOCKED_LINE_WIDTH = 3
DIVIDER_LINE_WIDTH = 4
DATA_MARKER_SIZE = 7


# =============================================================================
# 3. HELPER FUNCTIONS
# =============================================================================

def round_val(n, decimals=DECIMAL_PLACES):
    """Round to the configured number of decimal places."""
    factor = 10 ** decimals
    return np.round(n * factor) / factor


def calculate_median(arr):
    """Calculate median matching the TypeScript implementation."""
    if len(arr) == 0:
        return 0
    return float(np.median(arr))


# =============================================================================
# 4. DATA PREPARATION
# =============================================================================

def prepare_data(df):
    """
    Prepare the Power BI dataset for XmR charting.
    Expects a DataFrame with at least 2 columns: first is date, second is value.
    Returns (dates, values, measure_name) as numpy arrays + the column name.
    """
    cols = df.columns.tolist()
    date_col = cols[0]
    value_col = cols[1]

    # Drop rows with missing data
    df = df.dropna(subset=[date_col, value_col]).copy()

    # Parse dates
    df[date_col] = pd.to_datetime(df[date_col])

    # Ensure numeric values
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
    df = df.dropna(subset=[value_col])

    # Sort by date ascending
    df = df.sort_values(date_col).reset_index(drop=True)

    dates = df[date_col].values
    values = df[value_col].values.astype(float)

    return dates, values, value_col


# =============================================================================
# 5. STATISTICAL CALCULATIONS
# =============================================================================

def get_moving_range(values):
    """Calculate moving range: |x[i] - x[i-1]| for consecutive values."""
    if len(values) < 2:
        return np.array([])
    return np.abs(np.diff(values))


def calculate_limits(values):
    """
    Calculate XmR control limits for a segment of data.
    Returns dict with avgX, avgMR, UNPL, LNPL, URL, upperQuartile, lowerQuartile.
    """
    mr = get_moving_range(values)

    if USE_MEDIAN:
        avg_x = calculate_median(values)
        avg_mr = calculate_median(mr) if len(mr) > 0 else 0.0
        npl_scale = MEDIAN_NPL_SCALING
        url_scale = MEDIAN_URL_SCALING
    else:
        avg_x = float(np.mean(values))
        avg_mr = float(np.mean(mr)) if len(mr) > 0 else 0.0
        npl_scale = NPL_SCALING
        url_scale = URL_SCALING

    delta = npl_scale * avg_mr
    unpl = avg_x + delta
    lnpl = avg_x - delta
    url = url_scale * avg_mr
    upper_quartile = (unpl + avg_x) / 2
    lower_quartile = (lnpl + avg_x) / 2

    return {
        "avgX": round_val(avg_x),
        "avgMR": round_val(avg_mr),
        "UNPL": round_val(unpl),
        "LNPL": round_val(lnpl),
        "URL": round_val(url),
        "upperQuartile": round_val(upper_quartile),
        "lowerQuartile": round_val(lower_quartile),
    }


# =============================================================================
# 6. EXCEPTION DETECTION
# =============================================================================

def check_run_of_eight(statuses, values, avg):
    """
    Mark points where 8 consecutive values are on the same side of the center line.
    avg can be a scalar or an array (for trend lines).
    """
    n = len(values)
    if n < 8:
        return

    # Determine the avg value per point
    if np.isscalar(avg):
        avg_arr = np.full(n, avg)
    else:
        avg_arr = np.asarray(avg)

    # Use an 8-bit sliding window (matches TypeScript bit manipulation)
    above_or_below = 0
    for i in range(7):
        if values[i] > avg_arr[i]:
            above_or_below |= (1 << (i % 8))

    for i in range(7, n):
        if values[i] > avg_arr[i]:
            above_or_below |= (1 << (i % 8))
        else:
            above_or_below &= ~(1 << (i % 8))

        if above_or_below == 0 or above_or_below == 255:
            for j in range(i - 7, i + 1):
                statuses[j] = RUN_OF_EIGHT


def check_near_limit(statuses, values, lower_quartile, upper_quartile):
    """
    Mark points where 3 out of 4 consecutive values are beyond a quartile line.
    lower_quartile and upper_quartile can be scalars or arrays (for trend lines).
    """
    n = len(values)
    if n < 4:
        return

    if np.isscalar(lower_quartile):
        lq_arr = np.full(n, lower_quartile)
    else:
        lq_arr = np.asarray(lower_quartile)

    if np.isscalar(upper_quartile):
        uq_arr = np.full(n, upper_quartile)
    else:
        uq_arr = np.asarray(upper_quartile)

    below_count = 0
    above_count = 0

    # Initialize sliding window for first 3 points
    for i in range(3):
        if values[i] < lq_arr[i]:
            below_count += 1
        elif values[i] > uq_arr[i]:
            above_count += 1

    # Slide window across remaining points
    for i in range(3, n):
        if values[i] < lq_arr[i]:
            below_count += 1
        elif values[i] > uq_arr[i]:
            above_count += 1

        if below_count >= 3 or above_count >= 3:
            for j in range(i - 3, i + 1):
                statuses[j] = NEAR_LIMIT

        # Remove the leaving element from the window
        if values[i - 3] < lq_arr[i - 3]:
            below_count -= 1
        elif values[i - 3] > uq_arr[i - 3]:
            above_count -= 1


def check_outside_limit(statuses, values, lnpl, unpl):
    """
    Mark points that fall outside the natural process limits.
    lnpl and unpl can be scalars or arrays (for trend lines).
    """
    n = len(values)
    if np.isscalar(lnpl):
        lnpl_arr = np.full(n, lnpl)
    else:
        lnpl_arr = np.asarray(lnpl)

    if np.isscalar(unpl):
        unpl_arr = np.full(n, unpl)
    else:
        unpl_arr = np.asarray(unpl)

    for i in range(n):
        if values[i] < lnpl_arr[i] or values[i] > unpl_arr[i]:
            statuses[i] = OUTSIDE_NPL


def detect_exceptions(values, avg, lower_quartile, upper_quartile, lnpl, unpl):
    """
    Apply all three exception rules in order (later rules override earlier ones):
    1. Run of 8 (blue)
    2. Near limit (orange)
    3. Outside NPL (red)
    Returns an array of status codes.
    """
    n = len(values)
    statuses = np.full(n, NORMAL, dtype=int)

    check_run_of_eight(statuses, values, avg)
    check_near_limit(statuses, values, lower_quartile, upper_quartile)
    check_outside_limit(statuses, values, lnpl, unpl)

    return statuses


# =============================================================================
# 7. TREND LINES (Linear Regression)
# =============================================================================

def linear_regression(values, dates):
    """
    Fit y = mx + c using normalized x-values (matching XMRIT's normalization).
    Returns (m, c) or None if regression is not possible.
    """
    n = len(values)
    if n < 2:
        return None

    # Normalize dates: use the gap between first two points as the base unit
    dates_numeric = np.array([d.astype("datetime64[ms]").astype(np.int64) for d in dates], dtype=float)
    first = dates_numeric[0]
    base = dates_numeric[1] - first
    if base == 0:
        return None

    x_norm = (dates_numeric - first) / base
    y = values

    sum_x = np.sum(x_norm)
    sum_y = np.sum(y)
    sum_xy = np.sum(x_norm * y)
    sum_x2 = np.sum(x_norm ** 2)

    denom = n * sum_x2 - sum_x * sum_x
    if denom == 0:
        return None

    m = (n * sum_xy - sum_x * sum_y) / denom
    c = (sum_y - m * sum_x) / n

    return m, c


def create_trend_lines(values, dates):
    """
    Create trend lines for the X chart: center line, UNPL, LNPL, quartiles.
    Returns dict of arrays or None if regression fails.
    """
    result = linear_regression(values, dates)
    if result is None:
        return None

    m, c = result
    mr = get_moving_range(values)
    avg_mr = float(np.mean(mr)) if len(mr) > 0 else 0.0

    n = len(values)
    centre = np.zeros(n)
    unpl = np.zeros(n)
    lnpl = np.zeros(n)
    upper_qtl = np.zeros(n)
    lower_qtl = np.zeros(n)

    for i in range(n):
        cl = i * m + c
        u = cl + avg_mr * NPL_SCALING
        l = cl - avg_mr * NPL_SCALING
        centre[i] = round_val(cl)
        unpl[i] = round_val(u)
        lnpl[i] = round_val(l)
        upper_qtl[i] = round_val((u + cl) / 2)
        lower_qtl[i] = round_val((l + cl) / 2)

    return {
        "centre": centre,
        "unpl": unpl,
        "lnpl": lnpl,
        "upperQtl": upper_qtl,
        "lowerQtl": lower_qtl,
    }


# =============================================================================
# 8. SEASONALITY ADJUSTMENT
# =============================================================================

def determine_periodicity(dates):
    """Determine the most common interval between data points."""
    if len(dates) < 2:
        return "day"

    deltas = np.diff(dates.astype("datetime64[D]").astype(np.int64))
    if len(deltas) == 0:
        return "day"

    # Find most common delta
    unique, counts = np.unique(deltas, return_counts=True)
    most_common = unique[np.argmax(counts)]

    if most_common < 7:
        return "day"
    elif most_common < 28:
        return "week"
    elif most_common < 90:
        return "month"
    elif most_common < 365:
        return "quarter"
    else:
        return "year"


def get_sub_period_index(date, period):
    """
    Get the sub-period index within a period for a given date.
    For period='year': returns the position within the year based on data periodicity.
    """
    dt = pd.Timestamp(date)
    if period == "year":
        return dt.dayofyear
    elif period == "quarter":
        # Day within the quarter
        quarter_start = dt - pd.tseries.offsets.QuarterBegin(startingMonth=1)
        return (dt - quarter_start).days
    elif period == "month":
        return dt.day
    elif period == "week":
        return dt.dayofweek
    return 0


def deseasonalize_data(dates, values, period):
    """
    Remove seasonal patterns from the data.
    Groups data by sub-period position within each period, calculates seasonal
    factors (sub-period average / overall average), then divides values by factors.

    Returns deseasonalized values array.
    """
    if period is None or len(values) < 2:
        return values.copy()

    df = pd.DataFrame({"date": pd.to_datetime(dates), "value": values})

    # Determine the sub-period grouping key based on the seasonal period
    if period == "year":
        # Group by month (or week) within each year
        periodicity = determine_periodicity(dates)
        if periodicity in ("day", "week"):
            df["sub_key"] = df["date"].dt.isocalendar().week.astype(int)
        else:
            df["sub_key"] = df["date"].dt.month
    elif period == "quarter":
        df["sub_key"] = df["date"].dt.month % 3
        df.loc[df["sub_key"] == 0, "sub_key"] = 3
    elif period == "month":
        df["sub_key"] = df["date"].dt.day
    elif period == "week":
        df["sub_key"] = df["date"].dt.dayofweek
    else:
        return values.copy()

    # Calculate seasonal factors
    sub_period_avg = df.groupby("sub_key")["value"].mean()
    overall_avg = values.mean()

    if overall_avg == 0:
        return values.copy()

    seasonal_factors = sub_period_avg / overall_avg
    seasonal_factors = seasonal_factors.replace(0, 1)  # prevent division by zero

    # Apply seasonal factors
    result = values.copy()
    for i, row in df.iterrows():
        sf = seasonal_factors.get(row["sub_key"], 1.0)
        if sf != 0 and not np.isnan(sf):
            result[i] = values[i] / sf

    return result


# =============================================================================
# 9. SEGMENT PROCESSING (Dividers)
# =============================================================================

def build_segments(dates, values):
    """
    Split data into segments based on DIVIDER_DATES.
    Returns a list of segment dicts, each containing:
      - indices: array of original indices
      - dates: array of dates
      - values: array of values
      - limits: dict from calculate_limits()
      - statuses: array of exception status codes
      - is_first: whether this is the leftmost segment
      - is_last: whether this is the rightmost segment
    """
    # Parse divider dates and sort
    divider_dts = sorted([np.datetime64(d) for d in DIVIDER_DATES])

    # Build boundaries: [start, divider1, divider2, ..., end]
    boundaries = []
    boundaries.append(dates[0])
    for d in divider_dts:
        boundaries.append(d)
    boundaries.append(dates[-1] + np.timedelta64(1, "D"))  # inclusive end

    segments = []
    for seg_idx in range(len(boundaries) - 1):
        left = boundaries[seg_idx]
        right = boundaries[seg_idx + 1]

        # Find indices in this segment
        if seg_idx == len(boundaries) - 2:
            # Last segment: inclusive on right
            mask = (dates >= left) & (dates <= right)
        else:
            mask = (dates >= left) & (dates < right)

        indices = np.where(mask)[0]
        if len(indices) == 0:
            continue

        seg_dates = dates[indices]
        seg_values = values[indices]
        limits = calculate_limits(seg_values)

        is_first = (seg_idx == 0)
        is_last = (seg_idx == len(boundaries) - 2)

        # Determine which limits to use for exception detection
        if is_first and _locked_limits_active():
            locked = _get_locked_limits()
            avg_for_check = locked["avgX"]
            lq_for_check = locked["lowerQuartile"]
            uq_for_check = locked["upperQuartile"]
            lnpl_for_check = locked["LNPL"]
            unpl_for_check = locked["UNPL"]
        elif is_first and SHOW_TREND_LINES:
            trend = create_trend_lines(seg_values, seg_dates)
            if trend is not None:
                avg_for_check = trend["centre"]
                lq_for_check = trend["lowerQtl"]
                uq_for_check = trend["upperQtl"]
                lnpl_for_check = trend["lnpl"]
                unpl_for_check = trend["unpl"]
            else:
                avg_for_check = limits["avgX"]
                lq_for_check = limits["lowerQuartile"]
                uq_for_check = limits["upperQuartile"]
                lnpl_for_check = limits["LNPL"]
                unpl_for_check = limits["UNPL"]
        else:
            avg_for_check = limits["avgX"]
            lq_for_check = limits["lowerQuartile"]
            uq_for_check = limits["upperQuartile"]
            lnpl_for_check = limits["LNPL"]
            unpl_for_check = limits["UNPL"]

        statuses = detect_exceptions(
            seg_values,
            avg_for_check,
            lq_for_check,
            uq_for_check,
            lnpl_for_check,
            unpl_for_check,
        )

        segments.append({
            "indices": indices,
            "dates": seg_dates,
            "values": seg_values,
            "limits": limits,
            "statuses": statuses,
            "is_first": is_first,
            "is_last": is_last,
        })

    return segments


# =============================================================================
# 10. LOCKED LIMITS
# =============================================================================

def _locked_limits_active():
    """Check if any locked limit is configured."""
    return (LOCKED_AVG is not None or
            LOCKED_UNPL is not None or
            LOCKED_LNPL is not None)


def _get_locked_limits():
    """
    Build locked limits dict. For any limit not explicitly locked,
    fall back to None (the caller should use the calculated value).
    """
    # We need calculated limits as a baseline
    # This will be computed from the full dataset (no dividers)
    # and then overridden by any user-specified locked values.
    return {
        "avgX": LOCKED_AVG,
        "UNPL": LOCKED_UNPL,
        "LNPL": LOCKED_LNPL,
        "lowerQuartile": ((LOCKED_LNPL or 0) + (LOCKED_AVG or 0)) / 2 if LOCKED_AVG is not None and LOCKED_LNPL is not None else None,
        "upperQuartile": ((LOCKED_UNPL or 0) + (LOCKED_AVG or 0)) / 2 if LOCKED_AVG is not None and LOCKED_UNPL is not None else None,
    }


def get_effective_locked_limits(calculated_limits):
    """
    Merge locked limits with calculated limits (locked values take precedence).
    Returns a fully populated limits dict and whether quartiles should be shown.
    """
    locked = _get_locked_limits()
    merged = calculated_limits.copy()

    avg_modified = False
    unpl_modified = False
    lnpl_modified = False

    if LOCKED_AVG is not None:
        merged["avgX"] = LOCKED_AVG
        avg_modified = True
    if LOCKED_UNPL is not None:
        merged["UNPL"] = LOCKED_UNPL
        unpl_modified = True
    if LOCKED_LNPL is not None:
        merged["LNPL"] = LOCKED_LNPL
        lnpl_modified = True

    # Recalculate quartiles from the (possibly overridden) limits
    merged["lowerQuartile"] = round_val((merged["LNPL"] + merged["avgX"]) / 2)
    merged["upperQuartile"] = round_val((merged["UNPL"] + merged["avgX"]) / 2)

    # Determine quartile visibility (matching shouldUseQuartile logic)
    use_upper = True
    use_lower = True

    if not avg_modified and not unpl_modified and not lnpl_modified:
        pass  # both quartiles visible
    elif (lnpl_modified and unpl_modified) or avg_modified:
        is_sym = abs(merged["UNPL"] + merged["LNPL"] - 2 * merged["avgX"]) < 0.001
        if not is_sym:
            use_upper = False
            use_lower = False
    elif unpl_modified:
        use_upper = False
    elif lnpl_modified:
        use_lower = False

    return merged, use_upper, use_lower


# =============================================================================
# 11. CHART RENDERING
# =============================================================================

def render_chart(dates, values, measure_name):
    """
    Render the XMRIT-styled X chart using matplotlib.
    """
    # --- Apply seasonality adjustment if configured ---
    if SEASONAL_PERIOD is not None:
        values = deseasonalize_data(dates, values, SEASONAL_PERIOD)
        title_prefix = "Deseasonalised "
    else:
        title_prefix = ""

    # --- Build segments (handles dividers) ---
    segments = build_segments(dates, values)

    if len(segments) == 0:
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.text(0.5, 0.5, "No data to display", ha="center", va="center",
                transform=ax.transAxes, fontsize=14)
        plt.show()
        return

    # --- Create figure ---
    fig, ax = plt.subplots(figsize=(14, 5.5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # --- Plot connecting data line (black) across ALL points ---
    ax.plot(dates, values, color=DATA_LINE_COLOR, linewidth=1, zorder=5)

    # --- Plot data points colored by exception status ---
    all_statuses = np.full(len(values), NORMAL, dtype=int)
    for seg in segments:
        all_statuses[seg["indices"]] = seg["statuses"]

    for status_code in [NORMAL, RUN_OF_EIGHT, NEAR_LIMIT, OUTSIDE_NPL]:
        mask = all_statuses == status_code
        if np.any(mask):
            ax.scatter(
                dates[mask],
                values[mask],
                color=POINT_COLORS[status_code],
                s=DATA_MARKER_SIZE ** 2,
                zorder=10,
                edgecolors="none",
                clip_on=False,
            )

    # --- Add value labels above each point ---
    y_range = np.max(values) - np.min(values) if np.max(values) != np.min(values) else 1
    label_offset = y_range * 0.025
    for i in range(len(values)):
        label_text = f"{round_val(values[i])}"
        ax.annotate(
            label_text,
            (dates[i], values[i]),
            textcoords="offset points",
            xytext=(0, 8),
            ha="center",
            va="bottom",
            fontsize=8,
            fontweight="bold",
            color=LABEL_COLORS[all_statuses[i]],
            zorder=15,
            clip_on=False,
        )

    # --- Track y-axis extremes for auto-scaling ---
    chart_y_min = np.min(values)
    chart_y_max = np.max(values)

    # --- Draw limit lines and trend lines per segment ---
    for seg in segments:
        seg_dates = seg["dates"]
        seg_limits = seg["limits"]
        is_first = seg["is_first"]
        is_last = seg["is_last"]

        # Determine x-extent for limit lines
        x_left = seg_dates[0]
        if is_last:
            # Extend slightly beyond the last data point
            x_right = seg_dates[-1] + np.timedelta64(1, "D")
        else:
            x_right = seg_dates[-1]

        line_x = [x_left, x_right]

        # --- Determine line style (locked vs normal) ---
        if is_first and _locked_limits_active():
            effective_limits, use_upper_q, use_lower_q = get_effective_locked_limits(seg_limits)
            lw = LOCKED_LINE_WIDTH
            ls = "solid"
        else:
            effective_limits = seg_limits
            use_upper_q = True
            use_lower_q = True
            lw = LIMIT_LINE_WIDTH
            ls = "dashed"

        # --- Skip static limit lines for first segment when trend lines active ---
        skip_static = (is_first and SHOW_TREND_LINES)
        # If trend lines are on and there are no dividers, skip all static lines
        skip_all_static = (SHOW_TREND_LINES and len(DIVIDER_DATES) == 0)

        if not skip_static and not skip_all_static:
            # Mean / center line (red)
            avg_val = effective_limits["avgX"]
            ax.plot(line_x, [avg_val, avg_val],
                    color=MEAN_COLOR, linewidth=lw, linestyle=ls, zorder=7)

            # UNPL (steelblue)
            unpl_val = effective_limits["UNPL"]
            ax.plot(line_x, [unpl_val, unpl_val],
                    color=LIMIT_COLOR, linewidth=lw, linestyle=ls, zorder=7)

            # LNPL (steelblue)
            lnpl_val = effective_limits["LNPL"]
            ax.plot(line_x, [lnpl_val, lnpl_val],
                    color=LIMIT_COLOR, linewidth=lw, linestyle=ls, zorder=7)

            # Quartile lines (grey, dashed, thin)
            if SHOW_QUARTILE_LINES:
                if use_upper_q:
                    uq_val = effective_limits["upperQuartile"]
                    ax.plot(line_x, [uq_val, uq_val],
                            color=QUARTILE_COLOR, linewidth=1, linestyle="dashed", zorder=6)
                if use_lower_q:
                    lq_val = effective_limits["lowerQuartile"]
                    ax.plot(line_x, [lq_val, lq_val],
                            color=QUARTILE_COLOR, linewidth=1, linestyle="dashed", zorder=6)

            # Add value labels at the right end of limit lines (last segment only)
            if is_last:
                label_x = x_right
                for lbl_val, lbl_color in [
                    (avg_val, MEAN_COLOR),
                    (unpl_val, LIMIT_COLOR),
                    (lnpl_val, LIMIT_COLOR),
                ]:
                    ax.annotate(
                        f"{round_val(lbl_val)}",
                        (label_x, lbl_val),
                        textcoords="offset points",
                        xytext=(5, 0),
                        ha="left",
                        va="center",
                        fontsize=11,
                        color="#000",
                        zorder=15,
                        clip_on=False,
                    )

            # Update y-axis range
            chart_y_min = min(chart_y_min, effective_limits["LNPL"])
            chart_y_max = max(chart_y_max, effective_limits["UNPL"])

        # --- Trend lines (first segment only, or all data if no dividers) ---
        if is_first and SHOW_TREND_LINES:
            trend = create_trend_lines(seg["values"], seg_dates)
            if trend is not None:
                # Center line (red, solid, thick)
                ax.plot(seg_dates, trend["centre"],
                        color=MEAN_COLOR, linewidth=3, linestyle="solid", zorder=8)
                # UNPL trend (steelblue, solid, thick)
                ax.plot(seg_dates, trend["unpl"],
                        color=LIMIT_COLOR, linewidth=3, linestyle="solid", zorder=8)
                # LNPL trend (steelblue, solid, thick)
                ax.plot(seg_dates, trend["lnpl"],
                        color=LIMIT_COLOR, linewidth=3, linestyle="solid", zorder=8)
                # Quartile trends (grey, dashed, thin)
                if SHOW_QUARTILE_LINES:
                    ax.plot(seg_dates, trend["upperQtl"],
                            color=QUARTILE_COLOR, linewidth=1, linestyle="dashed", zorder=6)
                    ax.plot(seg_dates, trend["lowerQtl"],
                            color=QUARTILE_COLOR, linewidth=1, linestyle="dashed", zorder=6)

                # Update y-axis range for trend lines
                chart_y_min = min(chart_y_min, np.min(trend["lnpl"]))
                chart_y_max = max(chart_y_max, np.max(trend["unpl"]))

    # --- Draw divider lines ---
    for div_date_str in DIVIDER_DATES:
        div_dt = np.datetime64(div_date_str)
        ax.axvline(
            x=div_dt,
            color=DIVIDER_COLOR,
            linewidth=DIVIDER_LINE_WIDTH,
            linestyle="solid",
            zorder=20,
        )

    # --- Configure axes ---
    # X-axis: date formatting
    # Use %#d on Windows, %-d on Linux/Mac for day without leading zero
    import platform
    day_fmt = "%#d" if platform.system() == "Windows" else "%-d"
    ax.xaxis.set_major_formatter(mdates.DateFormatter(f"{day_fmt} %b"))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax.tick_params(axis="x", labelsize=11, colors="#000")
    plt.setp(ax.get_xticklabels(), rotation=0, ha="center")

    # Y-axis: numeric, no grid, ~6 intervals
    y_padding = (chart_y_max - chart_y_min) * PADDING_FROM_EXTREMES
    ax.set_ylim(chart_y_min - y_padding, chart_y_max + y_padding)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6))
    ax.tick_params(axis="y", labelsize=11, colors="#000")

    # Remove grid lines
    ax.grid(False)

    # Axis line styling (black)
    for spine in ax.spines.values():
        spine.set_color("#000")
        spine.set_linewidth(0.8)

    # --- Title ---
    title_text = f"{title_prefix}X Plot"
    if measure_name.lower() != "value":
        title_text += f": {measure_name}"
    ax.set_title(title_text, fontsize=13, fontweight="normal", color="#000", pad=10)

    # --- Layout ---
    fig.tight_layout()
    plt.show()


# =============================================================================
# 12. MAIN EXECUTION
# =============================================================================

# Power BI provides data in a DataFrame called `dataset`.
# For standalone testing, uncomment the lines below:
# dataset = pd.DataFrame({
#     "Date": pd.date_range("2020-01-01", periods=16, freq="D"),
#     "Value": [5045, 4350, 4350, 3975, 4290, 4430, 4485, 4285,
#               3980, 3925, 3645, 3760, 3300, 3685, 3463, 5200],
# })

try:
    dataset  # type: ignore  # noqa: F821 - provided by Power BI at runtime
except NameError:
    raise RuntimeError(
        "No 'dataset' DataFrame found. "
        "This script is designed to run inside a Power BI Python visual. "
        "Uncomment the sample dataset above for standalone testing."
    )

dates, values, measure_name = prepare_data(dataset)  # type: ignore
render_chart(dates, values, measure_name)

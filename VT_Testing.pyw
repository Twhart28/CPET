# vt_dashboard.py
# UI dashboard for VT analysis with V-slope, LOESS trends, and VT1/VT2 markers.
# Adds Manual Edit mode with draggable VT lines and Excel-style draggable V-slope segments,
# with Save/Cancel/Undo/Redo and strict separation from Auto mode.

# --- Crash popup + log for double-click launches ---
import os, sys, traceback
LOG_PATH = os.path.join(os.path.dirname(__file__), "vt_dashboard_error.log")

def _excepthook(exc_type, exc, tb):
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("Uncaught exception:\n")
            f.write("".join(traceback.format_exception(exc_type, exc, tb)))
    except Exception:
        pass
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        r = _tk.Tk(); r.withdraw()
        _mb.showerror("VT Dashboard crashed", f"See error log:\n{LOG_PATH}")
        r.destroy()
    except Exception:
        pass
sys.excepthook = _excepthook

# Force Tk backend for matplotlib (avoids backend auto-detect surprises)
import matplotlib
matplotlib.use("TkAgg")
# --- end crash hook ---

import math
import copy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.ticker import MultipleLocator
from matplotlib.lines import Line2D
from matplotlib.patches import Circle
from matplotlib.legend import Legend   # add

# Global drag lock so only one thing moves at a time
_DRAG_LOCK = {"vslope": False, "vt": False}

# ======================== Config (mirrors your VBA constants) ========================

MARKER_SIZE = 12
AXIS_PAD_FRAC = 0.05

COL_BLUE        = "#377EB7"
COL_ORANGE      = "#FF7F0E"
COL_BLUE_DARK   = "#1E5F96"
COL_ORANGE_DARK = "#C85A0A"
COL_PURPLE      = "#800080"
COL_RED         = "#FF0000"
COL_GREEN       = "#008C46"

LOESS_SPAN_FRAC = 0.30
LOESS_MIN_NEIGHBORS = 9

BS_MIN_POINTS_PER_SIDE = 12
BS_MIN_XSPAN = 0.5
BS_Q_LO = 0.15
BS_Q_HI = 0.85
BS2_PRE_SLOPE_MIN = 0.85
BS2_PRE_SLOPE_MAX = 1.20
BS2_MID_SLOPE_MIN = 1.10

VSLOPE_GAP_POINTS = 1   # how many VO2 samples to skip after breakpoints for next green segment min in manual mode V-Slope chart

# Data layout (0-indexed; row 30 in Excel == index 29)
FIRST_DATA_ROW_IDX = 29
COL_TIME   = 0   # A
COL_VO2    = 1   # B
COL_VCO2   = 4   # E
COL_VE     = 5   # F
COL_RER    = 6   # G
COL_PetO2  = 15  # P
COL_PetCO2 = 16  # Q

# chart scaling configs
BASE_FIG_W, BASE_FIG_H = 7.8, 4.8
CHART_SCALE_MIN, CHART_SCALE_MAX = 0.6, 2.2
CHART_SCALE_STEP = 1.08
MIN_FIG_HEIGHT_PX = 260

# Handle visuals (manual mode)
HANDLE_RADIUS = 0.06          # in data units (scaled per-axes)
HANDLE_PICKR = 8              # pixel pick radius
LINE_PICKR   = 5              # pixel pick radius

# ============================= Small math utilities =================================

def _nice_unit(vrng):
    if vrng <= 0:
        return 1.0
    exp = math.floor(math.log10(vrng))
    frac = vrng / (10 ** exp)
    if frac <= 1.2: base = 0.2
    elif frac <= 2.5: base = 0.5
    elif frac <= 5: base = 1.0
    elif frac <= 8: base = 2.0
    else: base = 5.0
    return base * (10 ** exp)

# ============================= LOESS (local linear) =================================

def loess(x, y, span_frac=LOESS_SPAN_FRAC, min_neighbors=LOESS_MIN_NEIGHBORS):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    n = len(x)
    out = np.full(n, np.nan)

    mask = np.isfinite(x) & np.isfinite(y)
    xv = x[mask]; yv = y[mask]
    m = len(xv)
    if m == 0:
        return out

    k = max(min_neighbors, int(math.ceil(span_frac * m)))
    k = min(k, m)

    for i in range(n):
        if not (np.isfinite(x[i]) and np.isfinite(y[i])):
            continue
        xi = x[i]
        d = np.abs(xv - xi)
        h = np.partition(d, k-1)[k-1]
        if h <= 0:
            out[i] = y[i]
            continue
        w = (1 - (d / h) ** 3) ** 3
        use = w > 0
        if not np.any(use):
            out[i] = y[i]
            continue
        w = w[use]
        X = np.column_stack([np.ones(np.sum(use)), xv[use]])
        Y = yv[use]
        W = np.diag(w)
        try:
            beta = np.linalg.lstsq(W @ X, W @ Y, rcond=None)[0]
            out[i] = beta[0] + beta[1] * xi
        except Exception:
            out[i] = y[i]
    return out

# ======================== VT detectors (ported patterns) =============================

def detect_min_time(t, y, q_lo=0.05, q_hi=0.95):
    t = np.asarray(t, float); y = np.asarray(y, float)
    n = len(t)
    mask = np.isfinite(t) & np.isfinite(y)
    if mask.sum() < 3:
        return np.nan
    idx = np.where(mask)[0]
    i_lo = max(idx[0], int(math.ceil(q_lo * n)))
    i_hi = min(idx[-1], int(math.floor(q_hi * n)))
    if i_lo >= i_hi:
        i_lo, i_hi = idx[0], idx[-1]
    seg = np.arange(i_lo, i_hi + 1, dtype=int)
    seg = seg[np.isfinite(y[seg])]
    if len(seg) == 0:
        return np.nan
    i_min = seg[np.argmin(y[seg])]
    return t[i_min]

def moving_slope(t, y, k):
    t = np.asarray(t, float); y = np.asarray(y, float)
    n = len(t)
    out = np.full(n, np.nan)
    for i in range(n):
        L = max(0, i - k); U = min(n - 1, i + k)
        tx = t[L:U+1]; yy = y[L:U+1]
        m = np.isfinite(tx) & np.isfinite(yy)
        tx = tx[m]; yy = yy[m]
        cnt = len(tx)
        if cnt >= 2:
            Sx = tx.sum(); Sy = yy.sum()
            Sxx = (tx * tx).sum(); Sxy = (tx * yy).sum()
            denom = cnt * Sxx - Sx * Sx
            if abs(denom) > 1e-12:
                out[i] = (cnt * Sxy - Sx * Sy) / denom
    return out

def detect_vevco2_steepest_rise_start_strict(t, y_smooth, slope_window_pts=12,
                                              pos_slope_threshold=0.02, min_run=8):
    t = np.asarray(t, float); y = np.asarray(y_smooth, float)
    n = len(t)
    if n < 5: return np.nan
    mask = np.isfinite(y)
    if mask.sum() == 0: return np.nan
    i_min = np.where(mask)[0][np.argmin(y[mask])]
    if i_min >= n - 2: return np.nan
    s = moving_slope(t, y, slope_window_pts // 2)
    tail = np.arange(i_min + 1, n)
    tail = tail[np.isfinite(s[tail])]
    if len(tail) == 0: return np.nan
    j_star = tail[np.argmax(s[tail])]
    if s[j_star] <= 0: return np.nan
    k = j_star
    while k > i_min + 1 and np.isfinite(s[k]) and s[k] >= pos_slope_threshold:
        k -= 1
    start_idx = min(k + 1, n - 1)
    run = 0
    for j in range(start_idx, n):
        if np.isfinite(s[j]) and s[j] >= pos_slope_threshold:
            run += 1
            if run >= min_run:
                break
        else:
            break
    if run < min_run: return np.nan
    return t[start_idx]

def detect_vt2_petco2_steepest_run_start(t, y_smooth, slope_window_pts=12,
                                         min_run=10, neg_slope_threshold=-0.03,
                                         min_cum_drop=0.0):
    t = np.asarray(t, float); y = np.asarray(y_smooth, float)
    n = len(t)
    if n < 5: return np.nan
    mask = np.isfinite(y)
    if mask.sum() == 0: return np.nan
    i_max = np.where(mask)[0][np.argmax(y[mask])]
    if i_max >= n - 2: return np.nan
    s = moving_slope(t, y, slope_window_pts // 2)
    best_i = -1; best_avg = np.inf
    for i in range(i_max + 1, n - min_run):
        block = s[i:i+min_run]
        if np.all(np.isfinite(block)):
            avg = np.mean(block)
            if min_cum_drop > 0 and (np.isfinite(y[i]) and np.isfinite(y[i+min_run])):
                if (y[i] - y[i+min_run]) < min_cum_drop:
                    continue
            if avg < best_avg:
                best_avg = avg
                best_i = i
    if best_i < 0: return np.nan
    k = best_i
    while k > i_max + 1 and np.isfinite(s[k]) and s[k] <= neg_slope_threshold:
        k -= 1
    start_idx = min(k + 1, n - 1)
    return t[start_idx]

# =========================== Broken-stick (2-break) ==================================

def _hinge(x, t):
    h = x - t
    h[h < 0] = 0.0
    return h

def broken_stick_fit2(x, y,
                      min_span=BS_MIN_XSPAN,
                      min_pts=BS_MIN_POINTS_PER_SIDE,
                      q_lo=BS_Q_LO, q_hi=BS_Q_HI,
                      pre_min=BS2_PRE_SLOPE_MIN,
                      pre_max=BS2_PRE_SLOPE_MAX,
                      mid_min=BS2_MID_SLOPE_MIN):
    x = np.asarray(x, float); y = np.asarray(y, float)
    mask = np.isfinite(x) & np.isfinite(y) & (x > 0) & (y > 0)
    xv = x[mask]; yv = y[mask]
    m = len(xv)
    if m < 3 * min_pts:
        return None
    order = np.argsort(xv)
    xv = xv[order]; yv = yv[order]
    m = len(xv)
    k_lo = max(min_pts, int(math.ceil(q_lo * m)))
    k_hi = min(m - min_pts, int(math.floor(q_hi * m)))
    if k_lo >= k_hi:
        return None

    best = dict(sse=np.inf)
    for k1 in range(k_lo, k_hi - min_pts + 1):
        t1 = xv[k1]
        if (t1 - xv[0]) < min_span:
            continue
        for k2 in range(k1 + min_pts, k_hi + 1):
            t2 = xv[k2]
            if (t2 - t1) < min_span:       # mid span
                continue
            if (xv[-1] - t2) < min_span:   # post span
                continue

            H1 = _hinge(xv, t1)
            H2 = _hinge(xv, t2)
            A = np.column_stack([np.ones_like(xv), xv, H1, H2])
            try:
                beta, *_ = np.linalg.lstsq(A, yv, rcond=None)
            except Exception:
                continue
            b0, b1, b2, b3 = beta
            s1 = b1
            s2 = b1 + b2
            s3 = b1 + b2 + b3
            if not (pre_min <= s1 <= pre_max):
                continue
            if not (s2 >= mid_min):
                continue
            yhat = A @ beta
            sse = float(np.sum((yv - yhat) ** 2))
            if sse < best['sse']:
                best = dict(sse=sse, t1=float(t1), t2=float(t2),
                            b0=float(b0), b1=float(b1), b2=float(b2), b3=float(b3),
                            s1=float(s1), s2=float(s2), s3=float(s3))
    return None if not np.isfinite(best.get('sse', np.inf)) else best

# ============================== Parsing Excel ========================================

def _sniff_excel_format(path: str) -> str:
    import os
    with open(path, 'rb') as f:
        head = f.read(8)
    if head.startswith(b'\xD0\xCF\x11\xE0'):
        return 'xls'
    if head.startswith(b'PK\x03\x04'):
        ext = os.path.splitext(path)[1].lower()
        if ext == '.xlsb':
            return 'xlsb'
        return 'xlsx'
    return 'unknown'

def load_input_excel(path):
    import xlrd
    import pandas as pd
    import numpy as np
    import os

    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        data = []
        for r in range(sheet.nrows):
            row = []
            for c in range(sheet.ncols):
                row.append(sheet.cell_value(r, c))
            data.append(row)
        df = pd.DataFrame(data)
    else:
        df = pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")

    if df.shape[0] <= FIRST_DATA_ROW_IDX:
        raise RuntimeError("File has fewer than 30 rows; cannot find data starting at row 30.")

    blk = df.iloc[FIRST_DATA_ROW_IDX:, :].replace("", np.nan)
    blk = blk.infer_objects(copy=False)

    def colsafe(i):
        if i >= blk.shape[1]:
            return np.full(len(blk), np.nan)
        return pd.to_numeric(blk.iloc[:, i], errors="coerce").to_numpy(dtype=float)

    t      = colsafe(COL_TIME)
    vo2    = colsafe(COL_VO2)
    vco2   = colsafe(COL_VCO2)
    ve     = colsafe(COL_VE)
    rer    = colsafe(COL_RER)
    peto2  = colsafe(COL_PetO2)
    petco2 = colsafe(COL_PetCO2)

    valid = np.isfinite(t)
    if valid.sum() == 0:
        raise RuntimeError("No valid time values found under column A starting at row 30.")
    last = np.where(valid)[0][-1]
    sl = slice(0, last + 1)

    return dict(
        t=t[sl], vo2=vo2[sl], vco2=vco2[sl], ve=ve[sl], rer=rer[sl],
        peto2=peto2[sl], petco2=petco2[sl]
    )

# ============================== Computation ==========================================

def compute_all(raw):
    t = raw["t"]; vo2 = raw["vo2"]; vco2 = raw["vco2"]; ve = raw["ve"]
    rer = raw["rer"]; peto2 = raw["peto2"]; petco2 = raw["petco2"]

    ve_vo2  = np.where(np.isfinite(vo2) & (vo2 != 0), ve / vo2, np.nan)
    ve_vco2 = np.where(np.isfinite(vco2) & (vco2 != 0), ve / vco2, np.nan)

    m_peto2  = np.nanmean(peto2)  if np.isfinite(peto2).any()  else np.nan
    m_petco2 = np.nanmean(petco2) if np.isfinite(petco2).any() else np.nan
    peto2n  = peto2 / m_peto2  if np.isfinite(m_peto2)  and m_peto2  != 0 else np.full_like(peto2, np.nan)
    petco2n = petco2 / m_petco2 if np.isfinite(m_petco2) and m_petco2 != 0 else np.full_like(petco2, np.nan)

    vo2_L     = loess(t, vo2, span_frac=0.55, min_neighbors=25)
    ve_vo2_L  = loess(t, ve_vo2)
    ve_vco2_L = loess(t, ve_vco2)
    peto2n_L  = loess(t, peto2n)
    petco2n_L = loess(t, petco2n)
    rer_L     = loess(t, rer)

    vt1_vevo2_t  = detect_min_time(t, ve_vo2_L, q_lo=0.05, q_hi=0.95)
    vt2_vevco2_t = detect_vevco2_steepest_rise_start_strict(t, ve_vco2_L, slope_window_pts=12,
                                                            pos_slope_threshold=0.02, min_run=10)
    vt1_peto2_t  = detect_min_time(t, peto2n_L, q_lo=0.05, q_hi=0.95)
    vt2_petco2_t = detect_vt2_petco2_steepest_run_start(t, petco2n_L, slope_window_pts=12,
                                                        min_run=10, neg_slope_threshold=-0.03, min_cum_drop=0.0)

    bs = broken_stick_fit2(vo2, vco2, min_span=BS_MIN_XSPAN, min_pts=BS_MIN_POINTS_PER_SIDE,
                           q_lo=BS_Q_LO, q_hi=BS_Q_HI,
                           pre_min=BS2_PRE_SLOPE_MIN, pre_max=BS2_PRE_SLOPE_MAX, mid_min=BS2_MID_SLOPE_MIN)

    def interp_at_x(x, y, x0):
        x = np.asarray(x, float); y = np.asarray(y, float)
        if not np.isfinite(x0): return np.nan
        for i in range(len(x) - 1):
            if np.isfinite(y[i]) and np.isfinite(y[i+1]):
                if (x[i] <= x0 <= x[i+1]) or (x[i] >= x0 >= x[i+1]):
                    if x[i+1] != x[i]:
                        return float(y[i] + (y[i+1]-y[i])*(x0 - x[i])/(x[i+1]-x[i]))
                    else:
                        return float(y[i])
        return np.nan

    vt1_vslope_vo2 = bs["t1"] if bs else np.nan
    vt2_vslope_vo2 = bs["t2"] if bs else np.nan

    vt1_vevo2_vo2  = interp_at_x(t, vo2_L, vt1_vevo2_t)
    vt2_vevco2_vo2 = interp_at_x(t, vo2_L, vt2_vevco2_t)
    vt1_peto2_vo2  = interp_at_x(t, vo2_L, vt1_peto2_t)
    vt2_petco2_vo2 = interp_at_x(t, vo2_L, vt2_petco2_t)

    def safe_mean(vals):
        arr = np.asarray(vals, float)
        arr = arr[np.isfinite(arr)]
        return float(arr.mean()) if arr.size else np.nan

    avg_vt1 = safe_mean([vt1_vslope_vo2, vt1_vevo2_vo2, vt1_peto2_vo2])
    avg_vt2 = safe_mean([vt2_vslope_vo2, vt2_vevco2_vo2, vt2_petco2_vo2])

    def invert_y_to_x(x, y, y0):
        x = np.asarray(x, float); y = np.asarray(y, float)
        if not np.isfinite(y0): return np.nan
        for i in range(len(x) - 1):
            if np.isfinite(y[i]) and np.isfinite(y[i+1]):
                yi, yj = y[i], y[i+1]
                if (yi <= y0 <= yj) or (yi >= y0 >= yj):
                    if yj != yi:
                        return float(x[i] + (x[i+1]-x[i])*(y0 - yi)/(yj - yi))
                    else:
                        return float(x[i])
        return np.nan

    t_at_avg_vt1 = invert_y_to_x(t, vo2_L, avg_vt1)
    t_at_avg_vt2 = invert_y_to_x(t, vo2_L, avg_vt2)
    rer_at_vt1   = interp_at_x(t, rer_L, t_at_avg_vt1)
    rer_at_vt2   = interp_at_x(t, rer_L, t_at_avg_vt2)

    out_df = pd.DataFrame({
        "Time_min": t, "VO2": vo2, "VCO2": vco2, "VE": ve, "RER": rer,
        "PetO2": peto2, "PetCO2": petco2, "VE/VO2": ve_vo2, "VE/VCO2": ve_vco2,
        "PetO2_norm": peto2n, "PetCO2_norm": petco2n,
        "VE/VO2_LOESS": ve_vo2_L, "VE/VCO2_LOESS": ve_vco2_L,
        "PetO2_norm_LOESS": peto2n_L, "PetCO2_norm_LOESS": petco2n_L,
        "RER_LOESS": rer_L, "VO2_LOESS": vo2_L,
    })

    summary = dict(
        vslope_t1=vt1_vslope_vo2, vslope_t2=vt2_vslope_vo2,
        vevo2_min_t=vt1_vevo2_t,  vevco2_rise_t=vt2_vevco2_t,
        peto2_min_t=vt1_peto2_t,  petco2_drop_t=vt2_petco2_t,
        avg_vt1=avg_vt1, avg_vt2=avg_vt2,
        rer_at_vt1=rer_at_vt1, rer_at_vt2=rer_at_vt2,
        bs=bs
    )

    return out_df, summary

# ============================== Plot helpers =========================================

def _pad_axes(ax):
    xmn, xmx = ax.get_xlim()
    ymn, ymx = ax.get_ylim()
    xpad = (xmx - xmn) * AXIS_PAD_FRAC
    ypad = (ymx - ymn) * AXIS_PAD_FRAC
    ax.set_xlim(xmn - xpad, xmx + xpad)
    ax.set_ylim(ymn - ypad, ymx + ypad)

    xr = (ax.get_xlim()[1] - ax.get_xlim()[0])
    yr = (ax.get_ylim()[1] - ax.get_ylim()[0])
    step_x = _nice_unit(max(xr, 1e-9))
    step_y = _nice_unit(max(yr, 1e-9))
    try:
        ax.xaxis.set_major_locator(MultipleLocator(step_x))
        ax.yaxis.set_major_locator(MultipleLocator(step_y))
        ax.xaxis.set_minor_locator(MultipleLocator(step_x/2))
        ax.yaxis.set_minor_locator(MultipleLocator(step_y/2))
    except Exception:
        pass

    ax.grid(True, which="major", alpha=0.35)
    ax.minorticks_on()
    ax.grid(True, which="minor", alpha=0.15)

# ============================== Manual Edit Helpers ==================================

def _clamp(val, lo, hi):
    if lo > hi: lo, hi = hi, lo
    return max(lo, min(hi, val))

def _vslope_default_segs(bs, vo2_array):
    """Pre/mid/post segments with a gap only on the MIN side after each breakpoint:
       seg1: [.., t1] ; seg2: [t1 + GAP, t2] ; seg3: [t2 + GAP, ..].
       Endpoints snap to existing VO2 samples."""
    if not bs or vo2_array is None:
        return None

    xs = np.asarray(vo2_array, float)
    xs = np.sort(np.unique(xs[np.isfinite(xs)]))
    if xs.size < 2:
        return None

    n = xs.size
    b0,b1,b2,b3 = bs["b0"], bs["b1"], bs["b2"], bs["b3"]
    t1,t2 = float(bs["t1"]), float(bs["t2"])
    g = int(VSLOPE_GAP_POINTS)

    def f_pre(x):  return b0 + b1*x
    def f_mid(x):  return b0 + b1*x + b2*max(0, x - t1)
    def f_post(x): return b0 + b1*x + b2*max(0, x - t1) + b3*max(0, x - t2)

    # Indices around t1 and t2 in VO2 sample grid
    i1 = int(np.searchsorted(xs, t1, side="left"))
    i2 = int(np.searchsorted(xs, t2, side="left"))

    # --- Segment 1: from xs[0] .. sample at/before t1 (no gap on max)
    pre_start = 0
    pre_end = min(n - 1, i1)
    if pre_end < n and xs[pre_end] > t1 and pre_end > 0:
        pre_end -= 1

    # --- Segment 2: from (t1 + GAP samples) .. sample at/before t2
    mid_start = min(n - 1, i1 + g)
    mid_end = min(n - 1, i2)
    if mid_end < n and xs[mid_end] > t2 and mid_end > mid_start:
        mid_end -= 1
    mid_start = min(mid_start, mid_end)  # guard if few points

    # --- Segment 3: from (t2 + GAP samples) .. last sample
    post_start = min(n - 1, i2 + g)
    post_end = n - 1
    post_start = min(post_start, post_end)

    segs = []
    if pre_end > pre_start:
        x0, x1 = xs[pre_start], xs[pre_end]
        segs.append(((x0, f_pre(x0)), (x1, f_pre(x1))))
    if mid_end > mid_start:
        x0, x1 = xs[mid_start], xs[mid_end]
        segs.append(((x0, f_mid(x0)), (x1, f_mid(x1))))
    if post_end > post_start:
        x0, x1 = xs[post_start], xs[post_end]
        segs.append(((x0, f_post(x0)), (x1, f_post(x1))))
    return segs

class ManualState:
    """Stores manual-mode geometry."""
    def __init__(self, vt_map=None, vslope_segs=None):
        # vt_map: dict like {"vslope": (x1,x2), "ratio": (t1,t2), "pets": (t1,t2), "rer": (t1,t2), "vo2t": (t1,t2)}
        self.vt = copy.deepcopy(vt_map) if vt_map else {}
        self.vslope_segs = copy.deepcopy(vslope_segs) if vslope_segs else None

    def clone(self):
        return ManualState(self.vt, self.vslope_segs)

    def equals(self, other) -> bool:
        if other is None: return False
        if sorted(self.vt.keys()) != sorted(other.vt.keys()):
            return False
        for k in self.vt:
            a = self.vt[k]; b = other.vt.get(k, (np.nan, np.nan))
            for aa, bb in zip(a, b):
                if (np.isnan(aa) and np.isnan(bb)): continue
                if not np.isfinite(aa) or not np.isfinite(bb) or abs(aa-bb) > 1e-9:
                    return False
        if bool(self.vslope_segs) != bool(other.vslope_segs):
            return False
        if not self.vslope_segs: return True
        for s1, s2 in zip(self.vslope_segs, other.vslope_segs):
            for p1, p2 in zip(s1, s2):
                if abs(p1[0]-p2[0])>1e-9 or abs(p1[1]-p2[1])>1e-9:
                    return False
        return True

# ----- VT line group (shared across axes, draggable x only on allowed axes) -----
# ----- VT lines for ONE axes (independent; optional dragging) -----
class VTLinesSingleAx:
    def __init__(self, ax, canvas, t1, t2, color1, color2, draggable, on_commit_move,yield_to_handles=None, vt_pickradius=6):
        self.ax = ax
        self.canvas = canvas
        self.color1 = color1
        self.color2 = color2
        self._yield_to_handles = yield_to_handles  # callable returning list of handle pairs
        self._vt_pickradius = vt_pickradius
        self.draggable = bool(draggable)
        self.on_commit_move = on_commit_move
        self.t1 = float(t1) if np.isfinite(t1) else np.nan
        self.t2 = float(t2) if np.isfinite(t2) else np.nan
        self.l1 = None
        self.l2 = None
        self._drag = dict(active=False, which=None)
        self._cids = []
        self._build()

    def _near_any_handle_px(self, mouse_event, tol_px=HANDLE_PICKR):
        """Return True if mouse is within tol_px of any V-slope handle on this axes."""
        if not callable(self._yield_to_handles):
            return False
        handles = self._yield_to_handles() or []
        if not handles:
            return False
        mx, my = mouse_event.x, mouse_event.y  # pixel coords
        # handles is list of (h1, h2)
        for h1, h2 in handles:
            for h in (h1, h2):
                # each handle is a Line2D with single data point
                xd = float(h.get_xdata()[0]); yd = float(h.get_ydata()[0])
                hx, hy = self.ax.transData.transform((xd, yd))
                dx = mx - hx; dy = my - hy
                if (dx*dx + dy*dy) ** 0.5 <= tol_px:
                    return True
        return False

    def _mk_line(self, x, color, label):
        if not np.isfinite(x):
            return None
        ln = self.ax.axvline(
            x, linestyle="--", color=color, lw=1.8, label=label, zorder=5,
            picker=True, pickradius=self._vt_pickradius
        )
        return ln

    def _build(self):
        # remove old
        for ln in (self.l1, self.l2):
            try: ln.remove()
            except Exception: pass
        self.l1 = self._mk_line(self.t1, self.color1, "VT1")
        self.l2 = self._mk_line(self.t2, self.color2, "VT2")

        # event connections (pick/drag only if draggable)
        for cid in self._cids:
            try: self.canvas.mpl_disconnect(cid)
            except Exception: pass
        self._cids.clear()
        if self.draggable:
            self._cids.append(self.canvas.mpl_connect("pick_event", self._on_pick))
            self._cids.append(self.canvas.mpl_connect("motion_notify_event", self._on_motion))
            self._cids.append(self.canvas.mpl_connect("button_release_event", self._on_release))
        self.canvas.draw_idle()

    def _on_pick(self, ev):
        if _DRAG_LOCK.get("vslope"):   # someone is dragging the green segments
            return
        # If cursor is effectively on a green handle, yield (let VSlopeSegments handle it)
        if self._near_any_handle_px(ev.mouseevent):
            return

        if ev.artist in (self.l1, self.l2):
            which = "vt1" if ev.artist is self.l1 else "vt2"
            self._drag.update(active=True, which=which)
            _DRAG_LOCK["vt"] = True

    def _on_motion(self, ev):
        if not self._drag.get("active") or ev.inaxes is not self.ax or ev.xdata is None:
            return
        if self._drag["which"] == "vt1" and np.isfinite(self.t1):
            self.t1 = float(ev.xdata)
            if self.l1: self.l1.set_xdata([self.t1, self.t1])
        elif self._drag["which"] == "vt2" and np.isfinite(self.t2):
            self.t2 = float(ev.xdata)
            if self.l2: self.l2.set_xdata([self.t2, self.t2])
        self.canvas.draw_idle()

    def _on_release(self, ev):
        if self._drag.get("active"):
            self._drag["active"] = False
            _DRAG_LOCK["vt"] = False
            if callable(self.on_commit_move):
                self.on_commit_move("vt")

    def set_times(self, t1, t2):
        self.t1 = float(t1) if np.isfinite(t1) else np.nan
        self.t2 = float(t2) if np.isfinite(t2) else np.nan
        self._build()

    def get_times(self):
        return (self.t1, self.t2)

    def set_visible(self, vis: bool):
        for ln in (self.l1, self.l2):
            if ln is not None:
                ln.set_visible(vis)

# ----- V-slope manual segments with draggable endpoints and whole-segment drag -----
class VSlopeSegments:
    def __init__(self, ax, canvas, segs, on_commit_move):
        """
        ax: axes to draw on
        segs: list of 3 segments [((x1,y1),(x2,y2)), ...] (can be fewer if missing)
        """
        self.ax = ax
        self.canvas = canvas
        self.on_commit_move = on_commit_move
        self.segs = copy.deepcopy(segs) if segs else []
        self.lines = []
        self.handles = []  # per seg: (h1, h2)
        self._drag = dict(active=False, mode=None, seg_idx=None, end=None, start_xy=None)

        self._build()

    def _build(self):
        # Remove any existing
        for ln in self.lines:
            try: ln.remove()
            except Exception: pass
        for pair in self.handles:
            for h in pair:
                try: h.remove()
                except Exception: pass
        self.lines.clear(); self.handles.clear()

        # build
        for (p1, p2) in self.segs:
            x = [p1[0], p2[0]]; y = [p1[1], p2[1]]
            ln = Line2D(x, y, color=COL_GREEN, lw=2.4, picker=True, pickradius=LINE_PICKR, zorder=6)
            self.ax.add_line(ln)
            # handles
            h1, = self.ax.plot([p1[0]], [p1[1]], marker="o", markersize=8,
                            markerfacecolor="white", markeredgecolor=COL_GREEN,
                            linestyle="None", zorder=7, picker=HANDLE_PICKR)
            h2, = self.ax.plot([p2[0]], [p2[1]], marker="o", markersize=8,
                            markerfacecolor="white", markeredgecolor=COL_GREEN,
                            linestyle="None", zorder=7, picker=HANDLE_PICKR)
            self.lines.append(ln)
            self.handles.append((h1, h2))

        self.cid_pick  = self.canvas.mpl_connect("pick_event", self._on_pick)
        self.cid_move  = self.canvas.mpl_connect("motion_notify_event", self._on_motion)
        self.cid_rel   = self.canvas.mpl_connect("button_release_event", self._on_release)

        self.canvas.draw_idle()

    def set_visible(self, vis: bool):
        for ln in self.lines:
            ln.set_visible(vis)
        for (h1,h2) in self.handles:
            h1.set_visible(vis); h2.set_visible(vis)

    def _on_pick(self, ev):
        if _DRAG_LOCK.get("vt"):   # VT line is being dragged
            return
        art = ev.artist
        # Prefer handles over lines
        for i, (h1, h2) in enumerate(self.handles):
            if art is h1:
                self._drag.update(active=True, mode="move_end", seg_idx=i, end=0)
                _DRAG_LOCK["vslope"] = True
                return
            if art is h2:
                self._drag.update(active=True, mode="move_end", seg_idx=i, end=1)
                _DRAG_LOCK["vslope"] = True
                return
        if art in self.lines:
            i = self.lines.index(art)
            self._drag.update(active=True, mode="move_seg", seg_idx=i,
                            start_xy=(ev.mouseevent.xdata, ev.mouseevent.ydata))
            _DRAG_LOCK["vslope"] = True
            return

    def _on_motion(self, ev):
        if not self._drag.get("active"):
            return
        if ev.inaxes is not self.ax or ev.xdata is None or ev.ydata is None:
            return

        mode = self._drag["mode"]; i = self._drag["seg_idx"]
        x0,x1 = self.ax.get_xlim(); y0,y1 = self.ax.get_ylim()

        if mode == "move_end":
            end = self._drag["end"]
            (p1, p2) = self.segs[i]
            if end == 0:
                p1 = (_clamp(ev.xdata, x0, x1), _clamp(ev.ydata, y0, y1))
            else:
                p2 = (_clamp(ev.xdata, x0, x1), _clamp(ev.ydata, y0, y1))
            self.segs[i] = (p1, p2)
        elif mode == "move_seg":
            sx,sy = self._drag["start_xy"]
            dx = (ev.xdata - sx); dy = (ev.ydata - sy)
            (p1, p2) = self.segs[i]
            p1 = (_clamp(p1[0]+dx, x0, x1), _clamp(p1[1]+dy, y0, y1))
            p2 = (_clamp(p2[0]+dx, x0, x1), _clamp(p2[1]+dy, y0, y1))
            self.segs[i] = (p1, p2)
            self._drag["start_xy"] = (ev.xdata, ev.ydata)

        # apply to artists
        (p1, p2) = self.segs[i]
        ln = self.lines[i]
        ln.set_xdata([p1[0], p2[0]]); ln.set_ydata([p1[1], p2[1]])
        h1,h2 = self.handles[i]
        h1.set_data([p1[0]], [p1[1]])
        h2.set_data([p2[0]], [p2[1]])
        self.canvas.draw_idle()

    def _on_release(self, ev):
        if self._drag.get("active"):
            self._drag["active"] = False
            _DRAG_LOCK["vslope"] = False
            if callable(self.on_commit_move):
                self.on_commit_move("vslope")

    def get_segments(self):
        return copy.deepcopy(self.segs)

    def set_segments(self, segs):
        self.segs = copy.deepcopy(segs) if segs else []
        self._build()

# ============================== Dashboard UI =========================================

class RawDataWindow(tk.Toplevel):
    def __init__(self, master, df: pd.DataFrame):
        super().__init__(master)
        self.title("Raw Data")
        self.geometry("1100x600")
        self.df = df
        top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=6)
        ttk.Button(top, text="Export CSV…", command=self.export_csv).pack(side="right")
        frame = ttk.Frame(self); frame.pack(fill="both", expand=True, padx=8, pady=6)
        self.tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        for c in df.columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=110, anchor="center")
        vals = df.replace(np.nan, "").values.tolist()
        for row in vals:
            self.tree.insert("", "end", values=row)

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Save CSV", defaultextension=".csv", filetypes=[("CSV", "*.csv")]
        )
        if not path: return
        try:
            self.df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

class ChartsPanel(ttk.Frame):
    def __init__(self, master, data_df: pd.DataFrame, summary: dict):
        super().__init__(master)
        self.data = data_df
        self.summary = summary
        self._cursor_xy = {}
        self.fig_scale = 1.0
        self._fig_canv = []

        self.show_loess = tk.BooleanVar(master=self, value=True)
        self.show_vt    = tk.BooleanVar(master=self, value=True)
        self.show_legend = tk.BooleanVar(master=self, value=True)
        self.manual_mode = tk.BooleanVar(master=self, value=False)

        # Manual-mode state + history
        self._auto_state = self._make_auto_state()
        self._manual_saved = None          # last saved ManualState
        self._manual_current = None        # current working ManualState
        self._undo = []                    # stack of ManualState
        self._redo = []                    # stack of ManualState
        self._dirty = False                # unsaved edits present

        # --- Controls row ---
        ctrl = ttk.Frame(self); ctrl.pack(fill="x", padx=10, pady=10)
        ttk.Button(ctrl, text="Raw Data", command=self.open_raw_data).pack(side="left")
        ttk.Checkbutton(ctrl, text="Show LOESS trendlines", variable=self.show_loess,
                        command=self.refresh_visibility).pack(side="left", padx=12)
        ttk.Checkbutton(ctrl, text="Show VT1/VT2 lines", variable=self.show_vt,
                        command=self.refresh_visibility).pack(side="left")
        ttk.Checkbutton(ctrl, text="Show legends", variable=self.show_legend,
                        command=self.refresh_visibility).pack(side="left", padx=12)

        # Manual toggle + action buttons
        ttk.Separator(ctrl, orient="vertical").pack(side="left", padx=10, fill="y")
        self.chk_manual = ttk.Checkbutton(ctrl, text="Manual edit", variable=self.manual_mode,
                                          command=self._toggle_manual_mode)
        self.chk_manual.pack(side="left", padx=8)

        self.btn_save   = ttk.Button(ctrl, text="Save",   command=self._save_manual)
        self.btn_cancel = ttk.Button(ctrl, text="Cancel", command=self._cancel_manual)
        self.btn_undo   = ttk.Button(ctrl, text="Undo ⌃Z", command=self._undo_cmd)
        self.btn_redo   = ttk.Button(ctrl, text="Redo ⌃Y", command=self._redo_cmd)
        # hidden until first edit
        for b in (self.btn_save, self.btn_cancel, self.btn_undo, self.btn_redo):
            b.pack_forget()

        # Bind shortcuts
        self.bind_all("<Control-z>", lambda e: self._undo_cmd())
        self.bind_all("<Control-y>", lambda e: self._redo_cmd())

        # --- Paned splitter ---
        self.pw = ttk.Panedwindow(self, orient="horizontal")
        self.pw.pack(fill="both", expand=True, padx=10, pady=(0,10))

        # ===== Left summary =====
        left_shell = ttk.Frame(self.pw)
        self.pw.add(left_shell, weight=1)

        self.left_canvas = tk.Canvas(left_shell, highlightthickness=0, borderwidth=0)
        lscroll = ttk.Scrollbar(left_shell, orient="vertical", command=self.left_canvas.yview)
        self.left_canvas.configure(yscrollcommand=lscroll.set)
        self.left_canvas.pack(side="left", fill="both", expand=True)
        lscroll.pack(side="right", fill="y")

        self.left_inner = ttk.Frame(self.left_canvas)
        self.left_canvas.create_window((0,0), window=self.left_inner, anchor="nw")
        self.left_inner.bind("<Configure>", lambda e: self.left_canvas.configure(
            scrollregion=self.left_canvas.bbox("all")))

        sumf = ttk.LabelFrame(self.left_inner, text="Summary")
        sumf.pack(fill="x", padx=8, pady=6)
        grid = [
            ("V-Slope VT1 (VO2)", "vslope_t1"),
            ("V-Slope VT2 (VO2)", "vslope_t2"),
            ("VE/VO2 min (t)", "vevo2_min_t"),
            ("VE/VCO2 rise start (t)", "vevco2_rise_t"),
            ("PetO2 min (t)", "peto2_min_t"),
            ("PetCO2 drop start (t)", "petco2_drop_t"),
            ("Average VT1 (VO2)", "avg_vt1"),
            ("Average VT2 (VO2)", "avg_vt2"),
            ("RER @ Avg VT1", "rer_at_vt1"),
            ("RER @ Avg VT2", "rer_at_vt2"),
        ]
        for r, (lab, key) in enumerate(grid):
            ttk.Label(sumf, text=lab).grid(row=r, column=0, sticky="w", padx=8, pady=2)
            val = self.summary.get(key, np.nan)
            s = "" if not np.isfinite(val) else f"{val:.2f}"
            ttk.Label(sumf, text=s).grid(row=r, column=1, sticky="w", padx=8, pady=2)

        # ===== Right charts (scrollable) =====
        right_shell = ttk.Frame(self.pw)
        self.pw.add(right_shell, weight=4)

        sc = ttk.Frame(right_shell)
        sc.grid(row=0, column=0, sticky="nsew")
        sc.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        right_shell.rowconfigure(0, weight=1)
        right_shell.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(sc, borderwidth=0, highlightthickness=0)
        yscroll = ttk.Scrollbar(sc, orient="vertical",   command=self.canvas.yview)
        xscroll = ttk.Scrollbar(sc, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        sc.rowconfigure(0, weight=1); sc.columnconfigure(0, weight=1)

        self.inner = ttk.Frame(self.canvas)
        self._inner_win_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))

        for c in range(2): self.inner.columnconfigure(c, weight=0)
        for r in range(3): self.inner.rowconfigure(r,  weight=0)

        self._wire_mousewheel_scope(
            self.left_canvas,
            on_scroll=lambda n: self._scroll_canvas_units(self.left_canvas, n),
            on_zoom=lambda s, ev=None: self._zoom_ui(s)
        )
        self._wire_mousewheel_scope(
            self.canvas,
            on_scroll=lambda n: self._scroll_canvas_units(self.canvas, n),
            on_zoom=lambda s, ev=None: self._zoom_charts_grid(s)
        )
        self._wire_mousewheel_scope(
            sc,
            on_scroll=lambda n: self._scroll_canvas_units(self.canvas, n),
            on_zoom=lambda s, ev=None: self._zoom_charts_grid(s)
        )
        self._wire_mousewheel_scope(
            self.inner,
            on_scroll=lambda n: self._scroll_canvas_units(self.canvas, n),
            on_zoom=lambda s, ev=None: self._zoom_charts_grid(s)
        )

        # --- visual sash indicator ---
        self._sash_line = tk.Frame(self, bg="black", width=2)
        def _place_sash(_e=None):
            try:
                x = self.pw.sashpos(0)
            except Exception:
                return
            px = self.pw.winfo_x()
            py = self.pw.winfo_y()
            self._sash_line.place(x=px + x, y=py, width=2, height=self.pw.winfo_height())
        self.after(60, _place_sash)
        self.pw.bind("<Configure>", _place_sash)
        self.pw.bind("<B1-Motion>", _place_sash)
        self.pw.bind("<ButtonRelease-1>", _place_sash)

        # finally build the charts
        self._build_charts()

    # Show the raw dataframe in a popup
    def open_raw_data(self):
        try:
            RawDataWindow(self, self.data)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- scrolling/zoom plumbing ----------------
    def _scroll_canvas_units(self, c: tk.Canvas, n: int):
        bbox = c.bbox("all")
        if not bbox: return
        content_h = bbox[3] - bbox[1]
        if content_h <= c.winfo_height(): return
        c.yview_scroll(n, "units")

    def _scroll_canvas_xunits(self, c: tk.Canvas, n: int):
        bbox = c.bbox("all")
        if not bbox: return
        content_w = bbox[2] - bbox[0]
        if content_w <= c.winfo_width(): return
        start, end = c.xview()
        if (n < 0 and start <= 0.0) or (n > 0 and end >= 1.0): return
        c.xview_scroll(n, "units")

    def _wire_mousewheel_scope(self, widget: tk.Widget, on_scroll, on_zoom=None):
        def _mw(e):
            delta = getattr(e, "delta", 0)
            if delta == 0: return "break"
            ctrl  = (e.state & 0x0004) != 0
            shift = (e.state & 0x0001) != 0
            if ctrl and on_zoom:
                on_zoom(1 if delta > 0 else -1, e)
            else:
                if shift:
                    self._scroll_canvas_xunits(self.canvas, -1 if delta > 0 else 1)
                else:
                    on_scroll(-1 if delta > 0 else 1)
            return "break"
        widget.bind("<MouseWheel>", _mw, add="+")
        def _mw_up(e):   e.delta = +120; return _mw(e)
        def _mw_down(e): e.delta = -120; return _mw(e)
        widget.bind("<Button-4>", _mw_up, add="+")
        widget.bind("<Button-5>", _mw_down, add="+")
        widget.bind("<Enter>", lambda e: widget.focus_set(), add="+")

    def _wire_figure_wheel(self, fig_widget: tk.Widget, ax, cv):
        def _mw(e):
            delta = getattr(e, "delta", 0)
            if delta == 0: return "break"
            ctrl  = (e.state & 0x0004) != 0
            shift = (e.state & 0x0001) != 0
            if ctrl:
                self._zoom_charts_grid(1 if delta > 0 else -1)
            else:
                if shift:
                    self._scroll_canvas_xunits(self.canvas, -1 if delta > 0 else 1)
                else:
                    self._scroll_canvas_units(self.canvas, -1 if delta > 0 else 1)
            return "break"
        fig_widget.bind("<MouseWheel>", _mw, add="+")
        def _mw_up(e):   e.delta = +120; return _mw(e)
        def _mw_down(e): e.delta = -120; return _mw(e)
        fig_widget.bind("<Button-4>", _mw_up, add="+")
        fig_widget.bind("<Button-5>", _mw_down, add="+")

    def _zoom_ui(self, sign):
        root = self.winfo_toplevel()
        try:
            cur = float(root.tk.call('tk', 'scaling'))
        except Exception:
            cur = 1.0
        factor = 1.05 if sign > 0 else (1/1.05)
        root.tk.call('tk', 'scaling', cur * factor)
        self.after(0, lambda: (
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")),
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        ))

    def _zoom_charts_grid(self, sign: int):
        new = self.fig_scale * (CHART_SCALE_STEP if sign > 0 else 1.0 / CHART_SCALE_STEP)
        new = min(max(new, CHART_SCALE_MIN), CHART_SCALE_MAX)
        if abs(new - self.fig_scale) < 1e-3:
            return
        self.fig_scale = new

        for fig, cv in self._fig_canv:
            fig.set_size_inches(BASE_FIG_W * self.fig_scale,
                                BASE_FIG_H * self.fig_scale, forward=True)
            wpx = int(fig.get_figwidth()  * fig.get_dpi())
            hpx = int(fig.get_figheight() * fig.get_dpi())
            if hpx < MIN_FIG_HEIGHT_PX:
                hpx = MIN_FIG_HEIGHT_PX
                fig.set_size_inches(wpx / fig.get_dpi(), hpx / fig.get_dpi(), forward=True)
            for ax in fig.axes:
                lab_fs  = max(7, int(11 * self.fig_scale))
                tick_fs = max(6, int(9  * self.fig_scale))
                ax.title.set_fontsize(lab_fs + 1)
                ax.xaxis.label.set_fontsize(lab_fs)
                ax.yaxis.label.set_fontsize(lab_fs)
                for t in ax.get_xticklabels() + ax.get_yticklabels():
                    t.set_fontsize(tick_fs)
            cv.get_tk_widget().config(width=wpx, height=hpx)
            cv.draw_idle()

        self.after(0, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    # ------------------------- Figure helpers -------------------------
    def _new_fig_ax(self, title, xlabel, ylabel):
        fig = Figure(figsize=(BASE_FIG_W * self.fig_scale, BASE_FIG_H * self.fig_scale), dpi=100)
        ax = fig.add_subplot(111)
        fig.subplots_adjust(left=0.11, right=0.98, bottom=0.16, top=0.90)
        ax.set_title(title); ax.set_xlabel(xlabel); ax.set_ylabel(ylabel)
        lab_fs  = max(7, int(11 * self.fig_scale))
        tick_fs = max(6, int(9  * self.fig_scale))
        ax.title.set_fontsize(lab_fs + 1)
        ax.xaxis.label.set_fontsize(lab_fs)
        ax.yaxis.label.set_fontsize(lab_fs)
        for t in ax.get_xticklabels() + ax.get_yticklabels():
            t.set_fontsize(tick_fs)
        return fig, ax

    def _add_canvas(self, fig, row, col):
        cv = FigureCanvasTkAgg(fig, master=self.inner)
        widget = cv.get_tk_widget()
        widget.grid(row=row, column=col, padx=6, pady=6, sticky="nw")
        cv.draw()
        self._fig_canv.append((fig, cv))
        return cv

    def _plot_vertical_auto(self, ax, x, label, color):
        return ax.axvline(x, linestyle="--", color=color, lw=1.5, label=label, zorder=3)

    def _attach_hover(self, fig, ax, canvas):
        ann = ax.annotate(
            "", xy=(0, 0), xytext=(12, 12), textcoords="offset points",
            bbox=dict(boxstyle="round", fc="w", ec="#555", alpha=0.95),
            annotation_clip=False, zorder=10_000
        )
        ann.set_visible(False); ann.set_clip_on(False)
        ann.set_ha("left"); ann.set_va("bottom")

        def on_move(ev):
            if ev.inaxes is ax and ev.xdata is not None and ev.ydata is not None:
                self._cursor_xy[ax] = (ev.xdata, ev.ydata)
                ann.xy = (ev.xdata, ev.ydata)
                ann.set_text(f"{ev.xdata:.2f}, {ev.ydata:.2f}")
                xpix, ypix = ax.transData.transform((ev.xdata, ev.ydata))
                axbox = ax.get_window_extent()
                pad = 8
                box_w, box_h = 110, 44
                ox, oy = 12, 12; ha, va = "left", "bottom"
                if xpix + box_w > axbox.x1 - pad:  ox = -12; ha = "right"
                if ypix + box_h > axbox.y1 - pad:  oy = -12; va = "top"
                if xpix - box_w < axbox.x0 + pad:  ox = 12;  ha = "left"
                if ypix - box_h < axbox.y0 + pad:  oy = 12;  va = "bottom"
                ann.xytext = (ox, oy); ann.set_ha(ha); ann.set_va(va)
                if not ann.get_visible(): ann.set_visible(True)
                canvas.draw_idle()
            elif ann.get_visible():
                ann.set_visible(False); canvas.draw_idle()
        fig.canvas.mpl_connect("motion_notify_event", on_move)

    # ------------------------- Build charts -------------------------
    def _build_charts(self):
        self._loess_artists = []
        self._legend_artists_auto = []
        self._legend_artists_manual = []
        self._vt_artists_auto = []
        self._vslope_auto_segments = []

        T = self.data["Time_min"].to_numpy(float)
        VO2 = self.data["VO2"].to_numpy(float)
        VCO2 = self.data["VCO2"].to_numpy(float)
        VE = self.data["VE"].to_numpy(float)
        RER = self.data["RER"].to_numpy(float)
        VE_VO2 = self.data["VE/VO2"].to_numpy(float)
        VE_VCO2 = self.data["VE/VCO2"].to_numpy(float)
        PetO2N = self.data["PetO2_norm"].to_numpy(float)
        PetCO2N = self.data["PetCO2_norm"].to_numpy(float)
        VE_VO2_L = self.data["VE/VO2_LOESS"].to_numpy(float)
        VE_VCO2_L = self.data["VE/VCO2_LOESS"].to_numpy(float)
        PetO2N_L  = self.data["PetO2_norm_LOESS"].to_numpy(float)
        PetCO2N_L = self.data["PetCO2_norm_LOESS"].to_numpy(float)
        RER_L = self.data["RER_LOESS"].to_numpy(float)
        VO2_L = self.data["VO2_LOESS"].to_numpy(float)

        bs = self.summary.get("bs")
        t_vt1 = self.summary.get("vevo2_min_t", np.nan)
        t_vt2 = self.summary.get("vevco2_rise_t", np.nan)

        # ---------- 1) VCO2 vs VO2 (V-slope) ----------
        fig1, ax1 = self._new_fig_ax("VCO2 vs VO2 (V-slope)", "VO2 (L/min)", "VCO2 (L/min)")
        ax1.scatter(VO2, VCO2, s=12, color=COL_BLUE, alpha=0.75, label="_nolegend_")
        if bs:
            x_sorted = np.sort(VO2[np.isfinite(VO2)]); xs = np.unique(x_sorted)
            y1 = np.full_like(xs, np.nan, dtype=float)
            y2 = np.full_like(xs, np.nan, dtype=float)
            y3 = np.full_like(xs, np.nan, dtype=float)
            b0,b1,b2,b3 = bs["b0"], bs["b1"], bs["b2"], bs["b3"]
            t1,t2 = bs["t1"], bs["t2"]
            for i,xv in enumerate(xs):
                if xv <= t1:      y1[i] = b0 + b1*xv
                elif xv <= t2:    y2[i] = b0 + b1*xv + b2*max(0, xv - t1)
                else:             y3[i] = b0 + b1*xv + b2*max(0, xv - t1) + b3*max(0, xv - t2)
            l1, = ax1.plot(xs, y1, color=COL_GREEN, lw=2,   label="_nolegend_")
            l2, = ax1.plot(xs, y2, color=COL_GREEN, lw=2,   label="_nolegend_")
            l3, = ax1.plot(xs, y3, color=COL_GREEN, lw=2.4, label="_nolegend_")
            self._vslope_auto_segments = [l1, l2, l3]
            v1 = self._plot_vertical_auto(ax1, bs["t1"], "VT1", COL_PURPLE); self._vt_artists_auto.append(v1)
            v2 = self._plot_vertical_auto(ax1, bs["t2"], "VT2", COL_RED);    self._vt_artists_auto.append(v2)
            leg = ax1.legend(handles=[v1, v2], loc="lower right"); leg.set_zorder(10)
            self._legend_artists_auto.append(leg)
        _pad_axes(ax1)
        self.cv1 = self._add_canvas(fig1, 0, 0)
        self._attach_hover(fig1, ax1, self.cv1)
        self._wire_figure_wheel(self.cv1.get_tk_widget(), ax1, self.cv1)

        # ---------- 2) VE/VO2 & VE/VCO2 vs Time ----------
        fig2, ax2 = self._new_fig_ax("VE/VO2 and VE/VCO2 vs Time", "Time (min)", "Ratio")
        sc1 = ax2.scatter(T, VE_VO2,  s=12, color=COL_BLUE,   alpha=0.75, label="VE/VO2")
        sc2 = ax2.scatter(T, VE_VCO2, s=12, color=COL_ORANGE, alpha=0.75, label="VE/VCO2")
        self._ratio_base_handles = [sc1, sc2]
        l1, = ax2.plot(T, VE_VO2_L,  lw=2, color=COL_BLUE_DARK,   label="_nolegend_")
        l2, = ax2.plot(T, VE_VCO2_L, lw=2, color=COL_ORANGE_DARK, label="_nolegend_")
        self._loess_artists += [l1, l2]
        handles = [sc1, sc2]
        if np.isfinite(t_vt1):
            v = self._plot_vertical_auto(ax2, t_vt1, "VT1", COL_PURPLE); self._vt_artists_auto.append(v); handles.append(v)
        if np.isfinite(t_vt2):
            v = self._plot_vertical_auto(ax2, t_vt2, "VT2", COL_RED);    self._vt_artists_auto.append(v); handles.append(v)
        leg = ax2.legend(handles=handles, loc="lower right"); leg.set_zorder(10)
        self._legend_artists_auto.append(leg)
        _pad_axes(ax2)
        self.cv2 = self._add_canvas(fig2, 0, 1)
        self._attach_hover(fig2, ax2, self.cv2)
        self._wire_figure_wheel(self.cv2.get_tk_widget(), ax2, self.cv2)

        # ---------- 3) PetO2_norm & PetCO2_norm vs Time ----------
        fig3, ax3 = self._new_fig_ax("PetO2 and PetCO2 (normalized) vs Time",
                                     "Time (min)", "Normalized (÷ series mean)")
        sc3 = ax3.scatter(T, PetO2N,  s=12, color=COL_BLUE,   alpha=0.75, label="PetO2")
        sc4 = ax3.scatter(T, PetCO2N, s=12, color=COL_ORANGE, alpha=0.75, label="PetCO2")
        self._pets_base_handles = [sc3, sc4]
        l3, = ax3.plot(T, PetO2N_L,  lw=2, color=COL_BLUE_DARK,   label="_nolegend_")
        l4, = ax3.plot(T, PetCO2N_L, lw=2, color=COL_ORANGE_DARK, label="_nolegend_")
        self._loess_artists += [l3, l4]
        handles = [sc3, sc4]
        if np.isfinite(t_vt1):
            v = self._plot_vertical_auto(ax3, t_vt1, "VT1", COL_PURPLE); self._vt_artists_auto.append(v); handles.append(v)
        if np.isfinite(t_vt2):
            v = self._plot_vertical_auto(ax3, t_vt2, "VT2", COL_RED);    self._vt_artists_auto.append(v); handles.append(v)
        leg = ax3.legend(handles=handles, loc="lower right"); leg.set_zorder(10)
        self._legend_artists_auto.append(leg)
        _pad_axes(ax3)
        self.cv3 = self._add_canvas(fig3, 1, 0)
        self._attach_hover(fig3, ax3, self.cv3)
        self._wire_figure_wheel(self.cv3.get_tk_widget(), ax3, self.cv3)

        # ---------- 4) RER vs Time (VT lines only; not draggable) ----------
        fig4, ax4 = self._new_fig_ax("RER vs Time", "Time (min)", "RER")
        ax4.scatter(T, RER, s=12, color=COL_BLUE, alpha=0.75, label="_nolegend_")
        l5, = ax4.plot(T, RER_L, lw=2, color=COL_BLUE_DARK, label="_nolegend_")
        self._loess_artists.append(l5)
        handles = []
        if np.isfinite(t_vt1):
            v = self._plot_vertical_auto(ax4, t_vt1, "VT1", COL_PURPLE); self._vt_artists_auto.append(v); handles.append(v)
        if np.isfinite(t_vt2):
            v = self._plot_vertical_auto(ax4, t_vt2, "VT2", COL_RED);    self._vt_artists_auto.append(v); handles.append(v)
        if handles:
            leg = ax4.legend(handles=handles, loc="lower right"); leg.set_zorder(10)
            self._legend_artists_auto.append(leg)
        _pad_axes(ax4)
        self.cv4 = self._add_canvas(fig4, 1, 1)
        self._attach_hover(fig4, ax4, self.cv4)
        self._wire_figure_wheel(self.cv4.get_tk_widget(), ax4, self.cv4)

        # ---------- 5) VO2 vs Time (VT lines only; not draggable) ----------
        fig5, ax5 = self._new_fig_ax("VO2 vs Time", "Time (min)", "VO2 (L/min)")
        ax5.scatter(T, VO2, s=12, color=COL_BLUE, alpha=0.75, label="_nolegend_")
        l6, = ax5.plot(T, VO2_L, lw=2, color=COL_BLUE_DARK, label="_nolegend_")
        self._loess_artists.append(l6)
        handles = []
        if np.isfinite(t_vt1):
            v = self._plot_vertical_auto(ax5, t_vt1, "VT1", COL_PURPLE); self._vt_artists_auto.append(v); handles.append(v)
        if np.isfinite(t_vt2):
            v = self._plot_vertical_auto(ax5, t_vt2, "VT2", COL_RED);    self._vt_artists_auto.append(v); handles.append(v)
        if handles:
            leg = ax5.legend(handles=handles, loc="lower right"); leg.set_zorder(10)
            self._legend_artists_auto.append(leg)
        _pad_axes(ax5)
        self.cv5 = self._add_canvas(fig5, 2, 0)
        self._attach_hover(fig5, ax5, self.cv5)
        self._wire_figure_wheel(self.cv5.get_tk_widget(), ax5, self.cv5)

        # references
        self._ax_vslope = ax1
        self._ax_rer    = ax4
        self._ax_vo2t   = ax5
        self._ax_time_ratio = ax2
        self._ax_time_pets  = ax3

        # Build manual layers (hidden by default)
        self._init_manual_layers()

        self.refresh_visibility()

    # ------------------------- Manual state / layers -------------------------
    def _make_auto_state(self) -> ManualState:
        bs = self.summary.get("bs")
        # Auto VT per chart:
        vt_map = dict()

        # 1) V-slope axis (x = VO2)
        if bs:
            vt_map["vslope"] = (float(bs["t1"]), float(bs["t2"]))
        else:
            vt_map["vslope"] = (np.nan, np.nan)

        # 2) VE/VO2 & VE/VCO2 vs time (use VE/VO2-min and VE/VCO2-rise start)
        vt_map["ratio"] = (
            float(self.summary.get("vevo2_min_t", np.nan)),
            float(self.summary.get("vevco2_rise_t", np.nan)),
        )

        # 3) PetO2/PetCO2 vs time — use the same times as “ratio”
        vt_map["pets"] = vt_map["ratio"]

        # 4) RER vs time (read-only)
        vt_map["rer"] = vt_map["ratio"]

        # 5) VO2 vs time (read-only)
        vt_map["vo2t"] = vt_map["ratio"]

        # Build default V-slope segments from broken-stick fit (auto-style, no gap)
        segs = None
        if bs:
            segs = _vslope_default_segs(bs, self.data["VO2"].to_numpy(float))
        return ManualState(vt_map, segs)

    def _init_manual_layers(self):
        # Initialize saved/current state from auto if first time
        if self._manual_saved is None and self._manual_current is None:
            base = self._make_auto_state()
            self._manual_saved = base.clone()
            self._manual_current = base.clone()

        vt = self._manual_current.vt

        # Manual V-slope segments
        self._vslope_manual = None
        if self._manual_current.vslope_segs:
            self._vslope_manual = VSlopeSegments(
                ax=self._ax_vslope,
                canvas=self.cv1.figure.canvas,
                segs=self._manual_current.vslope_segs,
                on_commit_move=self._on_manual_change
            )

        # Per-axis VT objects (draggable only on vslope/ratio/pets)
        self._vt_manual = {
            "vslope": VTLinesSingleAx(
                self._ax_vslope, self.cv1.figure.canvas,
                vt.get("vslope", (np.nan,np.nan))[0],
                vt.get("vslope", (np.nan,np.nan))[1],
                COL_PURPLE, COL_RED, True, self._on_manual_change,
                yield_to_handles=lambda: (self._vslope_manual.handles if getattr(self, "_vslope_manual", None) else []),
                vt_pickradius=max(3, HANDLE_PICKR - 3)),  # a bit smaller than handle radius
            "ratio":  VTLinesSingleAx(self._ax_time_ratio, self.cv2.figure.canvas,
                                    vt.get("ratio", (np.nan,np.nan))[0],
                                    vt.get("ratio", (np.nan,np.nan))[1],
                                    COL_PURPLE, COL_RED, True,  self._on_manual_change),
            "pets":   VTLinesSingleAx(self._ax_time_pets,  self.cv3.figure.canvas,
                                    vt.get("pets", (np.nan,np.nan))[0],
                                    vt.get("pets", (np.nan,np.nan))[1],
                                    COL_PURPLE, COL_RED, True,  self._on_manual_change),
            "rer":    VTLinesSingleAx(self._ax_rer,        self.cv4.figure.canvas,
                                    vt.get("rer", (np.nan,np.nan))[0],
                                    vt.get("rer", (np.nan,np.nan))[1],
                                    COL_PURPLE, COL_RED, False, self._on_manual_change),
            "vo2t":   VTLinesSingleAx(self._ax_vo2t,       self.cv5.figure.canvas,
                                    vt.get("vo2t", (np.nan,np.nan))[0],
                                    vt.get("vo2t", (np.nan,np.nan))[1],
                                    COL_PURPLE, COL_RED, False, self._on_manual_change),
        }

        # Build manual legends (include base series + VT lines), but don't overwrite auto legends
        self._legend_artists_manual = []
        for key, vtlines in self._vt_manual.items():
            ax = {"vslope": self._ax_vslope, "ratio": self._ax_time_ratio, "pets": self._ax_time_pets,
                "rer": self._ax_rer, "vo2t": self._ax_vo2t}[key]

            # base series per-axis
            base = []
            if key == "ratio":
                base = getattr(self, "_ratio_base_handles", [])
            elif key == "pets":
                base = getattr(self, "_pets_base_handles", [])

            handles = []
            handles.extend([h for h in base if h.get_label() != "_nolegend_"])
            if vtlines.l1 is not None: handles.append(vtlines.l1)
            if vtlines.l2 is not None: handles.append(vtlines.l2)
            if not handles:
                continue

            labels = [h.get_label() for h in handles]
            leg = Legend(ax, handles=handles, labels=labels, loc="lower right")
            ax.add_artist(leg)          # <-- coexist with auto legend
            leg.set_zorder(10)
            self._legend_artists_manual.append(leg)

        # Start hidden (auto mode default)
        self._set_manual_visible(False)

    def _set_manual_visible(self, vis: bool):
        # Auto layers
        for ln in self._vt_artists_auto:
            ln.set_visible((not vis) and self.show_vt.get())
        for ln in self._vslope_auto_segments:
            # Green V-slope lines = trendlines
            ln.set_visible((not vis) and self.show_loess.get())
        for leg in self._legend_artists_auto:
            leg.set_visible((not vis) and self.show_legend.get())

        # Manual layers
        if isinstance(self._vt_manual, dict):
            for vt in self._vt_manual.values():
                vt.set_visible(vis and self.show_vt.get())
        if getattr(self, "_vslope_manual", None):
            self._vslope_manual.set_visible(vis and self.show_loess.get())

        # Manual legends visibility is driven in refresh_visibility(),
        # but keep them consistent when flipping modes:
        for leg in getattr(self, "_legend_artists_manual", []):
            leg.set_visible(vis and self.show_legend.get())

        # redraw
        for _, cv in self._fig_canv:
            cv.draw_idle()

    def _toggle_manual_mode(self):
        want = self.manual_mode.get()
        if not want:
            # trying to leave manual mode
            if self._dirty:
                messagebox.showinfo("Unsaved edits",
                                    "You have unsaved manual edits.\nPress Save to keep them, or Cancel to discard.")
                # force it back on
                self.manual_mode.set(True)
                return
            # leaving manual mode -> show auto
            self._set_manual_visible(False)
            self._update_buttons()
            return

        # entering manual mode
        # if never initialized (e.g., after file load), set from auto
        if self._manual_current is None:
            self._manual_current = self._make_auto_state()
            self._manual_saved = self._manual_current.clone()

        # sync artists to current manual state
        self._apply_vt_map(self._manual_current.vt)
        if self._vslope_manual is None and self._manual_current.vslope_segs:
            self._vslope_manual = VSlopeSegments(
                ax=self._ax_vslope,
                canvas=self.cv1.figure.canvas,
                segs=self._manual_current.vslope_segs,
                on_commit_move=self._on_manual_change
            )
        elif self._vslope_manual and self._manual_current.vslope_segs:
            self._vslope_manual.set_segments(self._manual_current.vslope_segs)

        self._set_manual_visible(True)
        self._update_buttons()

    def _push_history(self):
        if self._manual_current is None: return
        self._undo.append(self._manual_current.clone())
        self._redo.clear()

    def _on_manual_change(self, _what):
        # capture new state after a drag completes
        if not self.manual_mode.get():
            return
        cur = self._collect_current_manual()
        if self._manual_current is None or not cur.equals(self._manual_current):
            self._push_history()
            self._manual_current = cur
            self._dirty = True
            self._update_buttons()

    def _collect_current_manual(self) -> ManualState:
        vt_map = {}
        for key, vtlines in (self._vt_manual or {}).items():
            vt_map[key] = vtlines.get_times()
        segs = self._vslope_manual.get_segments() if self._vslope_manual else None
        return ManualState(vt_map, segs)

    def _apply_vt_map(self, vt_map):
        if not vt_map: return
        for key, vtlines in (self._vt_manual or {}).items():
            if key in vt_map:
                t1, t2 = vt_map[key]
                vtlines.set_times(t1, t2)

    def _save_manual(self):
        if not self.manual_mode.get():
            return
        if self._manual_current is None:
            return
        self._manual_saved = self._manual_current.clone()
        self._dirty = False
        self._update_buttons()

    def _cancel_manual(self):
        if not self.manual_mode.get():
            return
        # Restore to last saved, or to auto if never saved
        base = self._manual_saved.clone() if self._manual_saved else self._auto_state.clone()
        self._manual_current = base.clone()
        # apply to artists
        if self._vt_manual:
            self._apply_vt_map(self._manual_current.vt)
        if base.vslope_segs:
            if self._vslope_manual:
                self._vslope_manual.set_segments(base.vslope_segs)
            else:
                self._vslope_manual = VSlopeSegments(
                    ax=self._ax_vslope,
                    canvas=self.cv1.figure.canvas,
                    segs=base.vslope_segs,
                    on_commit_move=self._on_manual_change
                )
        elif self._vslope_manual:
            # remove if none
            self._vslope_manual.set_segments([])
        self._undo.clear(); self._redo.clear()
        self._dirty = False
        self._update_buttons()
        # still in manual mode with restored state visible
        self._set_manual_visible(True)

    def _undo_cmd(self):
        if not self.manual_mode.get() or not self._undo:
            return
        last = self._undo.pop()
        self._redo.append(self._manual_current.clone())
        self._manual_current = last
        # apply
        if self._vt_manual:
            self._apply_vt_map(self._manual_current.vt)
        if last.vslope_segs:
            if self._vslope_manual:
                self._vslope_manual.set_segments(last.vslope_segs)
            else:
                self._vslope_manual = VSlopeSegments(
                    ax=self._ax_vslope,
                    canvas=self.cv1.figure.canvas,
                    segs=last.vslope_segs,
                    on_commit_move=self._on_manual_change
                )
        elif self._vslope_manual:
            self._vslope_manual.set_segments([])
        self._dirty = True
        self._update_buttons()

    def _redo_cmd(self):
        if not self.manual_mode.get() or not self._redo:
            return
        nxt = self._redo.pop()
        self._undo.append(self._manual_current.clone())
        self._manual_current = nxt
        # apply
        if self._vt_manual:
            self._apply_vt_map(self._manual_current.vt)
        if nxt.vslope_segs:
            if self._vslope_manual:
                self._vslope_manual.set_segments(nxt.vslope_segs)
            else:
                self._vslope_manual = VSlopeSegments(
                    ax=self._ax_vslope,
                    canvas=self.cv1.figure.canvas,
                    segs=nxt.vslope_segs,
                    on_commit_move=self._on_manual_change
                )
        elif self._vslope_manual:
            self._vslope_manual.set_segments([])
        self._dirty = True
        self._update_buttons()

    def _update_buttons(self):
        # show Save/Cancel/Undo/Redo only when manual mode AND at least one edit has been made
        # (per your spec: appear once an edit is made)
        if self.manual_mode.get() and self._dirty:
            # pack if hidden
            if not str(self.btn_save) in self.btn_save.master.children:
                # already in the same frame; ensure visible
                pass
            # bring them up (idempotent: tk ignores duplicates)
            self.btn_save.pack(side="left", padx=6)
            self.btn_cancel.pack(side="left", padx=2)
            self.btn_undo.pack(side="left", padx=6)
            self.btn_redo.pack(side="left", padx=2)
        else:
            # hide them
            for b in (self.btn_save, self.btn_cancel, self.btn_undo, self.btn_redo):
                b.pack_forget()

        # gray out undo/redo when empty
        self.btn_undo.state( ("disabled",) if not self._undo else ("!disabled",) )
        self.btn_redo.state( ("disabled",) if not self._redo else ("!disabled",) )

        # lock the manual toggle while there are unsaved edits
        self.chk_manual.state( ("disabled",) if (self.manual_mode.get() and self._dirty) else ("!disabled",) )

    # ------------------------- Visibility toggles -------------------------
    def refresh_visibility(self):
        # LOESS visibility
        for art in self._loess_artists:
            art.set_visible(self.show_loess.get())

        manual_on = self.manual_mode.get()

        if manual_on:
            # Hide auto artists
            for ln in self._vt_artists_auto:
                ln.set_visible(False)
            for ln in self._vslope_auto_segments:
                ln.set_visible(False)
            for leg in self._legend_artists_auto:
                leg.set_visible(False)

            # Show manual VT lines per axes
            for vt in (self._vt_manual or {}).values():
                vt.set_visible(self.show_vt.get())

            # Green V-slope segments follow trendlines toggle
            if self._vslope_manual:
                self._vslope_manual.set_visible(self.show_loess.get())

            # Manual legends follow legend toggle
            for leg in getattr(self, "_legend_artists_manual", []):
                leg.set_visible(self.show_legend.get())

        else:
            # Auto mode: VT and legends follow toggles
            for ln in self._vt_artists_auto:
                ln.set_visible(self.show_vt.get())
            for ln in self._vslope_auto_segments:
                ln.set_visible(self.show_loess.get())
            for leg in self._legend_artists_auto:
                leg.set_visible(self.show_legend.get())

            # Hide manual layers
            for vt in (self._vt_manual or {}).values():
                vt.set_visible(False)
            if self._vslope_manual:
                self._vslope_manual.set_visible(False)
            for leg in getattr(self, "_legend_artists_manual", []):
                leg.set_visible(False)

        # redraw
        for _, cv in (self._fig_canv or []):
            try: cv.draw_idle()
            except Exception: pass

# ------------------------- Welcome + main -------------------------

class Welcome(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VT Analysis")
        self.update_idletasks()
        self.geometry("460x240")
        self.resizable(False, False)
        outer = ttk.Frame(self, padding=20); outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Welcome to the VT Analysis Dashboard",
                  font=("Segoe UI", 14, "bold")).pack(pady=(10, 6))
        ttk.Label(outer, text="Choose an Excel file laid out like your VBA input sheet.",
                  font=("Segoe UI", 10)).pack(pady=(0, 16))
        ttk.Button(outer, text="Select File…", command=self.select_file).pack()
        ttk.Label(outer, text="Tip: .XLS needs xlrd==1.2.0; .XLSX uses openpyxl.",
                  foreground="#666").pack(pady=(16, 0))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Open Excel file",
            filetypes=[("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            raw = load_input_excel(path)
            data_df, summary = compute_all(raw)
        except Exception as e:
            messagebox.showerror("Error", f"Could not process file:\n{e}")
            return
        dash = tk.Toplevel(self)
        dash.title(f"VT Dashboard – {os.path.basename(path)}")
        dash.update_idletasks()
        _center_on_screen(dash, width=1400, height=950)
        ChartsPanel(dash, data_df, summary).pack(fill="both", expand=True)
        self.withdraw()

def _center_on_screen(win, width=None, height=None):
    if width and height:
        win.geometry(f"{width}x{height}")
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    w, h = win.winfo_width(), win.winfo_height()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

def main():
    try:
        app = Welcome()
        try:
            style = ttk.Style()
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass
        app.after(0, lambda: _center_on_screen(app))
        app.mainloop()
    except Exception:
        try:
            with open(LOG_PATH, "w", encoding="utf-8") as f:
                f.write("Uncaught exception in main():\n")
                f.write("".join(traceback.format_exception(*sys.exc_info())))
        except Exception:
            pass
        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            r = _tk.Tk(); r.withdraw()
            _mb.showerror("VT Dashboard crashed", f"See error log:\n{LOG_PATH}")
            r.destroy()
        except Exception:
            pass
        print(f"Fatal error. See log: {LOG_PATH}")
        try:
            input("Press Enter to close...")
        except Exception:
            pass

if __name__ == "__main__":
    main()

from __future__ import annotations

"""
PARAMETRIC SEGMENTED AIRFOIL GENERATOR - V10
=====================================================================

Changes relative to V9 (ANSYS mm import + smooth curves)
-------------------------------------------------------
A. No more micro-edges. Sub-micron edges at the channel junctions (e.g.
   0.0001 mm) are below ANSYS DesignModeler tolerance at mm scale (they
   only "worked" in m / cm because the model was then 1000x / 10x too
   big). Wall / surface vertices closer than JUNCTION_MIN_EDGE_MM to a
   junction are now dropped, and every exported edge is >= EXPORT_MIN_EDGE_MM.
B. Smooth export (EXPORT_SMOOTH): every exported branch is resampled by
   arc length with a SMOOTHLY varying spacing (fine where the curve bends
   tightly, coarse where it is flat, neighbour spacing ratio ~<=1.3), so a
   spline through the points does not wiggle.

Changes relative to V8 (kept from V9)
----------------------
1. densify_airfoil(): the loaded airfoil is rebuilt ONCE as a smooth,
   dense, cosine-spaced airfoil (passes exactly through every input point).
   This removes long straight facets caused by sparse Excel data
   (e.g. the straight run into the trailing edge).
2. Optional sharp trailing edge (TE_SHARP): the upper and lower surfaces
   fade together over the last TE_BLEND_PCT % of the chord and meet at
   one point, like a normal airfoil. Set TE_SHARP = False to keep the
   original blunt trailing edge. Set DENSIFY_INPUT = False to disable
   the whole rebuild and use the raw Excel points as they are.
3. Airfoil-surface branches are exported with a very small tolerance
   (EXPORT_SURFACE_TOLERANCE_MM, 0 = keep every vertex); channel walls
   use EXPORT_POINT_TOLERANCE_MM.

(V8 items kept: SurfaceFunction evaluates the polyline exactly, surface
pieces keep their vertices, reference line built from the surface vertices.)

Output rule
-----------
2 segments -> segment_1.csv, segment_2.csv
3 segments -> segment_1.csv, segment_2.csv, segment_3.csv
ANSYS DesignModeler TXT files are written separately (mm, XYPlane).
"""

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List
import csv
import json
import math
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import matplotlib

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from openpyxl import load_workbook
from scipy.interpolate import PchipInterpolator, BSpline
from scipy.optimize import brentq
from shapely.geometry import Polygon, LineString, MultiPolygon, Point


ANSYS_DECIMALS = 6
GEOMETRY_EPS = 1e-8

# ---- Input airfoil rebuild ------------------------------------------------
DENSIFY_INPUT = True        # smooth curve THROUGH every input point (fills gaps only)
DENSIFY_POINTS = 600        # points per surface (cosine spaced)
TE_SHARP = False            # False: keep the trailing edge exactly as in the coordinates
TE_BLEND_PCT = 8.0          # thickness is faded to zero over the last x % of chord

# ---- Export point reduction ----------------------------------------------
EXPORT_SMOOTH = True                  # arc-length resampling with smooth spacing (recommended)
EXPORT_MIN_EDGE_MM = 0.05             # no exported edge shorter than ~this (ANSYS mm-safe)
EXPORT_ANGLE_STEP_DEG = 1.5           # max turning angle between neighbouring exported points
EXPORT_SPACING_GRADATION = 0.25       # how fast spacing may change along a branch
JUNCTION_MIN_EDGE_MM = 0.05           # drop vertices closer than this to a channel junction
MICRO_MERGE_MM = 0.01                 # merge ring vertices closer than this (e.g. hair-thin trailing edge)
# Used only when EXPORT_SMOOTH = False (old Douglas-Peucker reduction):
EXPORT_POINT_TOLERANCE_MM = 0.005     # channel walls
EXPORT_SURFACE_TOLERANCE_MM = 0.0005  # airfoil surface (0 = keep all points)
EXPORT_CORNER_ANGLE_DEG = 12.0
EXPORT_MAX_SPAN_PCT = 2.0             # also the MAX spacing in smooth mode (% of chord)

# S-curve channel: angle (deg) by which the channel leaves / arrives at the
# airfoil surface, measured relative to the LOCAL surface tangent.
S_DEPARTURE_DEG = 18.0
# Automatic handle length = this fraction of the distance P0 -> P3.
S_AUTO_HANDLE_FRACTION = 0.40
# Spacing (mm) of the dense samples used to build the channel walls.
WALL_SAMPLE_MM = 0.02

CURVE_METHODS = ["Bezier", "B-spline", "S-Bezier"]


# ============================================================
# DATA CLASSES
# ============================================================

@dataclass
class ChannelSpec:
    name: str
    lower_xc_percent: float
    upper_xc_percent: float
    gap_mm: float
    bend_mm: float
    curve_method: str


@dataclass
class GeneratorConfig:
    chord_mm: float
    segments: int
    channels: List[ChannelSpec]


# ============================================================
# EXCEL INPUT
# ============================================================

def read_airfoil_excel(path: str | Path) -> np.ndarray:
    """
    Read an Excel file with:

        Section_ID | Point_ID | X_mm | Y_mm | Z_mm

    Preferred worksheet: Airfoil_Coordinates
    Only X_mm and Y_mm are used for 2-D geometry.
    """

    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Input file does not exist:\n{path}")

    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Please select an Excel .xlsx or .xlsm file.")

    wb = load_workbook(path, read_only=True, data_only=True)

    try:
        if "Airfoil_Coordinates" in wb.sheetnames:
            ws = wb["Airfoil_Coordinates"]
        else:
            ws = wb[wb.sheetnames[0]]

        rows = ws.iter_rows(values_only=True)

        try:
            header = next(rows)
        except StopIteration:
            raise ValueError("The selected worksheet is empty.")

        normalized = []
        for item in header:
            if item is None:
                normalized.append("")
            else:
                normalized.append(str(item).strip().lower().replace(" ", "_"))

        required = ["section_id", "point_id", "x_mm", "y_mm", "z_mm"]
        missing = [col for col in required if col not in normalized]

        if missing:
            raise ValueError(
                "The Excel file is missing:\n"
                + ", ".join(missing)
                + "\n\nRequired columns:\n"
                "Section_ID | Point_ID | X_mm | Y_mm | Z_mm"
            )

        x_col = normalized.index("x_mm")
        y_col = normalized.index("y_mm")

        data = []

        for row in rows:
            if row is None:
                continue
            if len(row) <= max(x_col, y_col):
                continue

            x = row[x_col]
            y = row[y_col]

            try:
                x = float(x)
                y = float(y)
            except (TypeError, ValueError):
                continue

            if math.isfinite(x) and math.isfinite(y):
                data.append((x, y))

    finally:
        wb.close()

    if len(data) < 20:
        raise ValueError(
            f"Only {len(data)} valid X-Y points were found. "
            "Check the Excel coordinate sheet."
        )

    points = np.asarray(data, dtype=float)

    # Remove consecutive duplicates.
    filtered = [points[0]]
    for point in points[1:]:
        if np.linalg.norm(point - filtered[-1]) > GEOMETRY_EPS:
            filtered.append(point)

    points = np.asarray(filtered, dtype=float)

    # Close internally for polygon operations.
    if np.linalg.norm(points[0] - points[-1]) > GEOMETRY_EPS:
        points = np.vstack([points, points[0]])

    return points


def scale_to_chord(points: np.ndarray, chord_mm: float) -> np.ndarray:
    """Scale X coordinates to exactly 0..chord_mm."""

    if chord_mm <= 0:
        raise ValueError("Chord must be greater than 0 mm.")

    points = np.asarray(points, dtype=float)

    xmin = float(np.min(points[:, 0]))
    xmax = float(np.max(points[:, 0]))

    original_chord = xmax - xmin

    if original_chord <= GEOMETRY_EPS:
        raise ValueError("Input airfoil has zero X extent.")

    scale = chord_mm / original_chord

    scaled = points.copy()
    scaled[:, 0] = (scaled[:, 0] - xmin) * scale
    scaled[:, 1] = scaled[:, 1] * scale

    if np.linalg.norm(scaled[0] - scaled[-1]) > GEOMETRY_EPS:
        scaled = np.vstack([scaled, scaled[0]])

    return scaled


# ============================================================
# AIRFOIL SURFACE RECONSTRUCTION
# ============================================================

def _prepare_branch(branch: np.ndarray, chord_mm: float):
    branch = np.asarray(branch, dtype=float)

    order = np.argsort(branch[:, 0])
    branch = branch[order]

    x = branch[:, 0]
    y = branch[:, 1]

    rounded_x = np.round(x, 9)

    unique_x = []
    unique_y = []

    for xr in np.unique(rounded_x):
        mask = np.isclose(rounded_x, xr, atol=1e-9)
        unique_x.append(float(np.mean(x[mask])))
        unique_y.append(float(np.mean(y[mask])))

    x = np.asarray(unique_x, dtype=float)
    y = np.asarray(unique_y, dtype=float)

    if len(x) < 5:
        raise ValueError("A surface has too few unique X coordinates.")

    x[0] = 0.0
    x[-1] = chord_mm

    return x, y


class SurfaceFunction:
    """
    y(x) of one airfoil surface.

    GEOMETRY: evaluated exactly on the polyline through the surface points
    (piecewise linear), so the airfoil is never altered by the segmenter.

    SLOPE: estimated from a PCHIP spline in u = sqrt(x) (smooth, finite
    near the leading edge). Used for the S-curve end directions and by
    densify_airfoil().
    """

    def __init__(self, x: np.ndarray, y: np.ndarray):
        self.x = np.asarray(x, dtype=float)
        self.y = np.asarray(y, dtype=float)
        self._p = PchipInterpolator(
            np.sqrt(np.maximum(self.x, 0.0)), self.y, extrapolate=False
        )
        self._dp = self._p.derivative()

    def __call__(self, x):
        return np.interp(
            np.asarray(x, dtype=float), self.x, self.y,
            left=np.nan, right=np.nan,
        )

    def smooth(self, x):
        """Smooth (spline) evaluation, passes through every input point."""
        u = np.sqrt(np.clip(np.asarray(x, dtype=float), 0.0, None))
        return self._p(u)

    def vertices_between(self, x0: float, x1: float) -> np.ndarray:
        """Surface vertices strictly between x0 and x1."""
        m = (self.x > x0 + 1e-9) & (self.x < x1 - 1e-9)
        return np.column_stack([self.x[m], self.y[m]])

    def slope(self, x):
        """dy/dx at x (finite-clamped near the leading edge)."""
        u = np.sqrt(np.clip(np.asarray(x, dtype=float), 0.0, None))
        u = np.maximum(u, 1e-3)
        return self._dp(u) / (2.0 * u)


def build_surface_functions(points: np.ndarray):
    """
    Build upper and lower y(x) functions from a standard closed airfoil
    coordinate order (supports the common TE -> LE -> TE ordering).
    """

    contour = np.asarray(points[:-1], dtype=float)

    le = int(np.argmin(contour[:, 0]))

    rotated = np.roll(contour, -le, axis=0)

    # Trailing-edge point(s): a blunt trailing edge has TWO points at
    # x = x_max (one per surface). Each surface must keep its own TE point,
    # otherwise they get averaged and the TE collapses to the mean height.
    x_max = float(np.max(rotated[:, 0]))
    tol = 1e-6 * max(x_max, 1.0)
    te_ids = np.flatnonzero(rotated[:, 0] >= x_max - tol)
    te_a = int(te_ids[0])
    te_b = int(te_ids[-1]) if len(te_ids) >= 2 else te_a

    if te_a < 2 or te_b > len(rotated) - 3:
        raise ValueError("Could not identify the upper/lower airfoil surfaces.")

    branch_1 = rotated[:te_a + 1]
    branch_2 = np.vstack([rotated[te_b:], rotated[0:1]])

    xu1, yu1 = _prepare_branch(branch_1, float(np.max(points[:, 0])))
    xu2, yu2 = _prepare_branch(branch_2, float(np.max(points[:, 0])))

    if np.mean(yu1) > np.mean(yu2):
        upper = (xu1, yu1)
        lower = (xu2, yu2)
    else:
        upper = (xu2, yu2)
        lower = (xu1, yu1)

    upper_fun = SurfaceFunction(upper[0], upper[1])
    lower_fun = SurfaceFunction(lower[0], lower[1])

    return upper_fun, lower_fun


def densify_airfoil(
    points: np.ndarray,
    n: int = DENSIFY_POINTS,
    sharp_te: bool = TE_SHARP,
    blend_pct: float = TE_BLEND_PCT,
) -> np.ndarray:
    """
    Rebuild the airfoil as a smooth, dense, cosine-spaced closed contour
    (lower surface LE -> TE, then upper surface TE -> LE, closed).

    The spline passes exactly through every input point, so it only fills
    the gaps between them (no long straight facets).

    If sharp_te is True, the thickness is faded to zero over the last
    blend_pct % of the chord (smoothstep), so both surfaces meet at the
    mean trailing-edge point instead of ending in a blunt vertical edge.
    """
    up, lo = build_surface_functions(points)
    c = float(lo.x[-1])

    xs = c * 0.5 * (1.0 - np.cos(np.linspace(0.0, np.pi, int(n))))
    xs[0] = 0.0
    xs[-1] = c

    yl = lo.smooth(xs)
    yu = up.smooth(xs)

    if not (np.all(np.isfinite(yl)) and np.all(np.isfinite(yu))):
        raise ValueError("Could not rebuild the airfoil from the input points.")

    # Both surfaces share the single leading-edge point.
    y_le = 0.5 * (yu[0] + yl[0])
    yu[0] = y_le
    yl[0] = y_le

    if sharp_te and blend_pct > 0.0:
        x0 = c * (1.0 - blend_pct / 100.0)
        t = np.clip((xs - x0) / (c - x0), 0.0, 1.0)
        w = t * t * (3.0 - 2.0 * t)          # smoothstep: no kink at x0
        ym = 0.5 * (yu + yl)
        yu = yu * (1.0 - w) + ym * w
        yl = yl * (1.0 - w) + ym * w
        yu[-1] = yl[-1] = 0.5 * (yu[-1] + yl[-1])

    lower = np.column_stack([xs, yl])
    upper = np.column_stack([xs, yu])

    up_rev = upper[::-1]
    if sharp_te and blend_pct > 0.0:
        up_rev = up_rev[1:]                  # drop duplicate TE point
    up_rev = up_rev[:-1]                     # drop duplicate LE point

    ring = np.vstack([lower, up_rev])
    return np.vstack([ring, ring[:1]])


def airfoil_surface_point(
    points: np.ndarray,
    chord_mm: float,
    percent_chord: float,
    surface: str,
) -> np.ndarray:
    """Return the airfoil point at a specified percentage of chord."""

    if not 0.0 <= percent_chord <= 100.0:
        raise ValueError(
            f"{surface.title()} location must be between 0 and 100 %c."
        )

    upper_fun, lower_fun = build_surface_functions(points)

    x = (percent_chord / 100.0) * chord_mm

    if surface.lower() == "upper":
        y = float(upper_fun(x))
    else:
        y = float(lower_fun(x))

    return np.array([x, y], dtype=float)


# ============================================================
# CURVE GENERATION
# ============================================================

def cubic_bezier(
    p0: np.ndarray,
    p1: np.ndarray,
    p2: np.ndarray,
    p3: np.ndarray,
    n: int = 400,
) -> np.ndarray:

    t = np.linspace(0.0, 1.0, n)
    omt = 1.0 - t

    return (
        (omt ** 3)[:, None] * p0
        + (3 * omt ** 2 * t)[:, None] * p1
        + (3 * omt * t ** 2)[:, None] * p2
        + (t ** 3)[:, None] * p3
    )


def cubic_bspline(control_points: np.ndarray, n: int = 400) -> np.ndarray:

    control_points = np.asarray(control_points, dtype=float)

    if len(control_points) != 6:
        raise ValueError("The cubic B-spline needs 6 control points.")

    degree = 3

    knots = np.array(
        [0, 0, 0, 0, 0.5, 0.5, 1, 1, 1, 1],
        dtype=float,
    )

    u = np.linspace(0.0, 1.0, n)

    bx = BSpline(knots, control_points[:, 0], degree)
    by = BSpline(knots, control_points[:, 1], degree)

    return np.column_stack([bx(u), by(u)])


def generate_channel_curve(
    lower_point: np.ndarray,
    upper_point: np.ndarray,
    bend_mm: float,
    method: str,
    n: int = 400,
    lower_slope: float | None = None,
    upper_slope: float | None = None,
) -> np.ndarray:

    p0 = np.asarray(lower_point, dtype=float)
    p3 = np.asarray(upper_point, dtype=float)

    direction = p3 - p0
    length = float(np.linalg.norm(direction))

    if length <= GEOMETRY_EPS:
        raise ValueError(
            "Channel lower and upper attachment points are coincident."
        )

    tangent = direction / length
    normal = np.array([-tangent[1], tangent[0]])

    method = method.strip().lower()

    # --------------------------------------------------------
    # S-shaped cubic Bezier: handles leave P0 and arrive at P3 with
    # tangents flatter than the straight line P0->P3 -> S shape.
    # --------------------------------------------------------
    if method in {"s-bezier", "s-curve", "sbezier"}:
        dx = abs(p3[0] - p0[0])
        dy = p3[1] - p0[1]
        sgn = 1.0 if dy >= 0 else -1.0

        # Handle length. Bend field > 0 overrides it (mm); 0 = automatic.
        h = bend_mm if bend_mm > 0 else S_AUTO_HANDLE_FRACTION * length

        base0 = math.atan(lower_slope) if lower_slope is not None else 0.0
        base1 = math.atan(upper_slope) if upper_slope is not None else 0.0
        delta = math.radians(S_DEPARTURE_DEG)
        a0 = base0 + sgn * delta
        a1 = base1 + sgn * delta

        phi = math.atan2(abs(dy), max(dx, GEOMETRY_EPS))
        lim = 0.7 * phi
        a0 = sgn * min(sgn * a0, lim)
        a1 = sgn * min(sgn * a1, lim)

        p1 = p0 + h * np.array([math.cos(a0), math.sin(a0)])
        p2 = p3 - h * np.array([math.cos(a1), math.sin(a1)])
        return cubic_bezier(p0, p1, p2, p3, n=n)

    # Positive bend moves the middle portion toward the channel normal.
    p1 = p0 + 0.33 * direction + bend_mm * normal
    p2 = p0 + 0.67 * direction + bend_mm * normal

    if method == "bezier":
        return cubic_bezier(p0, p1, p2, p3, n=n)

    if method in {"b-spline", "bspline"}:
        control = np.vstack(
            [
                p0,
                p0 + 0.20 * direction + bend_mm * normal,
                p0 + 0.40 * direction + bend_mm * normal,
                p0 + 0.60 * direction + bend_mm * normal,
                p0 + 0.80 * direction + bend_mm * normal,
                p3,
            ]
        )

        return cubic_bspline(control, n=n)

    raise ValueError(f"Unsupported curve method: {method}")


# ============================================================
# SEGMENTATION GEOMETRY
# ============================================================

def resample_surface_branch(
    surface_fun,
    x_start: float,
    x_end: float,
    n: int = 600,
    trim_start: bool = False,
    trim_end: bool = False,
) -> np.ndarray:
    """
    Return the airfoil surface between two X locations (ascending) as an
    ordered polyline made of the exact start point, every surface vertex
    strictly in between, and the exact end point. Nothing is resampled or
    smoothed here. (n is kept only for call compatibility.)
    """

    if x_end <= x_start + GEOMETRY_EPS:
        raise ValueError(
            "Invalid segment boundary: an attachment location is not "
            "further aft (larger X) than the one before it. Check the "
            "lower/upper attachment percentages for each channel."
        )

    p_start = np.array([[x_start, float(surface_fun(x_start))]])
    p_end = np.array([[x_end, float(surface_fun(x_end))]])
    inner = surface_fun.vertices_between(x_start, x_end)

    # Drop surface vertices that sit (almost) on a channel junction so no
    # micro-edge is created there. Not applied at the LE / TE ends.
    if len(inner):
        keep = np.ones(len(inner), dtype=bool)
        if trim_start:
            keep &= np.linalg.norm(inner - p_start[0], axis=1) > JUNCTION_MIN_EDGE_MM
        if trim_end:
            keep &= np.linalg.norm(inner - p_end[0], axis=1) > JUNCTION_MIN_EDGE_MM
        inner = inner[keep]

    piece = np.vstack([p_start, inner, p_end])

    if not np.all(np.isfinite(piece)):
        raise ValueError(
            "Could not evaluate the airfoil surface between "
            f"X = {x_start:.3f} mm and X = {x_end:.3f} mm. Check the "
            "attachment percentages against the airfoil geometry."
        )

    return piece


def _dense_extended_curve(curve: np.ndarray, chord_mm: float):
    """
    Resample the centreline densely by arc length and extend it past both
    ends along its end tangents, so each offset wall is guaranteed to cross
    the airfoil surface somewhere on the curve.
    """
    curve = np.asarray(curve, dtype=float)
    seg = np.linalg.norm(np.diff(curve, axis=0), axis=1)
    s = np.concatenate([[0.0], np.cumsum(seg)])
    total = float(s[-1])

    n_dense = int(min(max(total / WALL_SAMPLE_MM, 800), 8000))
    su = np.linspace(0.0, total, n_dense)
    dense = np.column_stack(
        [np.interp(su, s, curve[:, 0]), np.interp(su, s, curve[:, 1])]
    )
    ds = total / (n_dense - 1)

    k = min(8, n_dense - 1)
    t0 = dense[k] - dense[0]
    t0 /= max(np.linalg.norm(t0), GEOMETRY_EPS)
    t1 = dense[-1] - dense[-1 - k]
    t1 /= max(np.linalg.norm(t1), GEOMETRY_EPS)

    ext = min(0.3 * chord_mm, 60.0)
    n_ext = max(int(ext / ds), 2)
    steps = np.arange(n_ext, 0, -1)[:, None] * ds

    before = dense[0] - steps * t0
    after = dense[-1] + steps[::-1] * t1

    extended = np.vstack([before, dense, after])
    return extended, n_ext, n_ext + n_dense - 1


def _wall_surface_crossing(wall: np.ndarray, surface_fun, anchor_index: int):
    """
    Exact crossing of a (dense) wall polyline with a surface function,
    choosing the crossing closest to the channel attachment. The crossing
    is refined with a root solver on the surface function.
    """
    x_min, x_max = float(surface_fun.x[0]), float(surface_fun.x[-1])
    g = wall[:, 1] - surface_fun(np.clip(wall[:, 0], x_min, x_max))
    ok = np.isfinite(g)
    prod = g[:-1] * g[1:]
    cand = np.flatnonzero(ok[:-1] & ok[1:] & (prod < 0.0))

    if len(cand) == 0:
        raise ValueError(
            "Could not find where the channel wall meets the airfoil "
            "surface. Reduce the gap, or move this attachment away "
            "from the leading/trailing edge."
        )

    k = int(cand[np.argmin(np.abs(cand - anchor_index))])
    p_a, p_b = wall[k], wall[k + 1]

    def f(t):
        px = p_a[0] + t * (p_b[0] - p_a[0])
        py = p_a[1] + t * (p_b[1] - p_a[1])
        return py - float(surface_fun(min(max(px, x_min), x_max)))

    t = brentq(f, 0.0, 1.0, xtol=1e-14)
    x = float(np.clip(p_a[0] + t * (p_b[0] - p_a[0]), x_min, x_max))
    y = float(surface_fun(x))  # exactly ON the surface
    return k, x, y


def _trim_wall(wall: np.ndarray, c_lo, c_up) -> np.ndarray:
    """Wall polyline from the lower crossing to the upper crossing."""
    k_lo, x_lo, y_lo = c_lo
    k_up, x_up, y_up = c_up

    if k_up <= k_lo:
        raise ValueError(
            "The channel wall does not run from the lower to the upper "
            "surface. Increase the lower/upper attachment separation or "
            "reduce the gap."
        )

    inner = wall[k_lo + 1:k_up + 1]
    if len(inner):
        d0 = np.linalg.norm(inner - np.array([x_lo, y_lo]), axis=1)
        d1 = np.linalg.norm(inner - np.array([x_up, y_up]), axis=1)
        inner = inner[(d0 > JUNCTION_MIN_EDGE_MM) & (d1 > JUNCTION_MIN_EDGE_MM)]

    return np.vstack([[[x_lo, y_lo]], inner, [[x_up, y_up]]])


def build_channel_walls(
    lower_point: np.ndarray,
    upper_point: np.ndarray,
    curve: np.ndarray,
    gap_mm: float,
    lower_fun,
    upper_fun,
):
    """
    Build the two boundary walls of a constant-width channel.

    Each wall is the centreline offset by +/- gap/2 perpendicular to the
    LOCAL tangent, cut exactly where it crosses the lower and upper
    airfoil surfaces.

    Returns forward_wall, aft_wall, lower_x_forward, lower_x_aft,
    upper_x_forward, upper_x_aft (walls ordered lower -> upper).
    """

    chord = float(lower_fun.x[-1])
    ext_curve, i_lo, i_up = _dense_extended_curve(curve, chord)

    tangent = np.gradient(ext_curve, axis=0)
    norm = np.linalg.norm(tangent, axis=1, keepdims=True)
    norm[norm < GEOMETRY_EPS] = GEOMETRY_EPS
    tangent /= norm
    normal = np.column_stack([-tangent[:, 1], tangent[:, 0]])

    half = gap_mm / 2.0
    wall_plus = ext_curve + half * normal
    wall_minus = ext_curve - half * normal

    # Leading-edge side decided where the walls are well separated in X.
    mid = (i_lo + i_up) // 2
    forward_sign = 1.0 if wall_plus[mid, 0] <= wall_minus[mid, 0] else -1.0

    fwd = wall_plus if forward_sign > 0 else wall_minus
    aft = wall_minus if forward_sign > 0 else wall_plus

    f_lo = _wall_surface_crossing(fwd, lower_fun, i_lo)
    f_up = _wall_surface_crossing(fwd, upper_fun, i_up)
    a_lo = _wall_surface_crossing(aft, lower_fun, i_lo)
    a_up = _wall_surface_crossing(aft, upper_fun, i_up)

    if a_lo[1] <= f_lo[1] or a_up[1] <= f_up[1]:
        raise ValueError(
            "The two channel walls overlap on the airfoil surface. "
            "Reduce the gap or the bend."
        )

    forward_wall = _trim_wall(fwd, f_lo, f_up)
    aft_wall = _trim_wall(aft, a_lo, a_up)

    return forward_wall, aft_wall, f_lo[1], a_lo[1], f_up[1], a_up[1]


def generate_geometry(
    points: np.ndarray,
    config: GeneratorConfig,
):
    """
    Build every segment as one explicit, deliberately ordered closed
    boundary:

        airfoil surface -> channel wall -> airfoil surface -> channel wall
    """

    upper_fun, lower_fun = build_surface_functions(points)

    chord = config.chord_mm
    channels = list(config.channels)

    lower_xs = [chord * c.lower_xc_percent / 100.0 for c in channels]
    upper_xs = [chord * c.upper_xc_percent / 100.0 for c in channels]

    for i in range(1, len(channels)):
        if lower_xs[i] <= lower_xs[i - 1] + GEOMETRY_EPS:
            raise ValueError(
                f"{channels[i].name}: lower attachment "
                f"({channels[i].lower_xc_percent:.2f} %c) must be "
                f"further aft than {channels[i - 1].name}'s "
                f"({channels[i - 1].lower_xc_percent:.2f} %c)."
            )
        if upper_xs[i] <= upper_xs[i - 1] + GEOMETRY_EPS:
            raise ValueError(
                f"{channels[i].name}: upper attachment "
                f"({channels[i].upper_xc_percent:.2f} %c) must be "
                f"further aft than {channels[i - 1].name}'s "
                f"({channels[i - 1].upper_xc_percent:.2f} %c)."
            )

    channel_curves: List[np.ndarray] = []
    forward_walls: List[np.ndarray] = []
    aft_walls: List[np.ndarray] = []
    channel_slots: List[Polygon] = []
    lower_x_forward_list: List[float] = []
    lower_x_aft_list: List[float] = []
    upper_x_forward_list: List[float] = []
    upper_x_aft_list: List[float] = []

    for channel, lower_x, upper_x in zip(channels, lower_xs, upper_xs):

        if channel.gap_mm <= 0:
            raise ValueError(f"{channel.name}: gap must be greater than 0 mm.")

        lower_point = np.array([lower_x, float(lower_fun(lower_x))])
        upper_point = np.array([upper_x, float(upper_fun(upper_x))])

        curve = generate_channel_curve(
            lower_point,
            upper_point,
            channel.bend_mm,
            channel.curve_method,
            lower_slope=float(lower_fun.slope(lower_x)),
            upper_slope=float(upper_fun.slope(upper_x)),
        )

        try:
            (
                forward_wall,
                aft_wall,
                lower_x_forward,
                lower_x_aft,
                upper_x_forward,
                upper_x_aft,
            ) = build_channel_walls(
                lower_point,
                upper_point,
                curve,
                channel.gap_mm,
                lower_fun,
                upper_fun,
            )
        except ValueError as exc:
            raise ValueError(f"{channel.name}: {exc}") from exc

        channel_curves.append(curve)
        forward_walls.append(forward_wall)
        aft_walls.append(aft_wall)
        lower_x_forward_list.append(lower_x_forward)
        lower_x_aft_list.append(lower_x_aft)
        upper_x_forward_list.append(upper_x_forward)
        upper_x_aft_list.append(upper_x_aft)

        slot_ring = np.vstack(
            [forward_wall, aft_wall[::-1], forward_wall[:1]]
        )
        channel_slots.append(Polygon(slot_ring))

    n_channels = len(channels)
    segments: List[Polygon] = []

    for j in range(n_channels + 1):

        lower_x_start = 0.0 if j == 0 else lower_x_aft_list[j - 1]
        lower_x_end = lower_x_forward_list[j] if j < n_channels else chord

        upper_x_start = 0.0 if j == 0 else upper_x_aft_list[j - 1]
        upper_x_end = upper_x_forward_list[j] if j < n_channels else chord

        lower_piece = resample_surface_branch(
            lower_fun, lower_x_start, lower_x_end,
            trim_start=(j > 0), trim_end=(j < n_channels),
        )

        upper_piece = resample_surface_branch(
            upper_fun, upper_x_start, upper_x_end,
            trim_start=(j > 0), trim_end=(j < n_channels),
        )[::-1]

        boundary = [lower_piece]

        if j < n_channels:
            boundary.append(forward_walls[j])

        boundary.append(upper_piece)

        if j > 0:
            boundary.append(aft_walls[j - 1][::-1])

        loop = np.vstack(boundary)

        keep_mask = np.ones(len(loop), dtype=bool)
        keep_mask[1:] = np.linalg.norm(np.diff(loop, axis=0), axis=1) > 1e-9
        loop = loop[keep_mask]

        segment_polygon = Polygon(loop)

        if not segment_polygon.is_valid:
            # Cleanup only: resolves razor-thin numerical slivers.
            repaired = segment_polygon.buffer(0)

            if isinstance(repaired, MultiPolygon):
                repaired = max(repaired.geoms, key=lambda g: g.area)

            area_before = Polygon(loop).area
            if (
                isinstance(repaired, Polygon)
                and repaired.is_valid
                and repaired.area > 0
                and abs(repaired.area - area_before)
                < 0.02 * max(area_before, GEOMETRY_EPS)
            ):
                segment_polygon = repaired
            else:
                raise ValueError(
                    f"Segment {j + 1}: the channel geometry produces "
                    "an invalid or zero-area boundary. Reduce "
                    "gap/bend, or move the attachment points further "
                    "apart."
                )

        if segment_polygon.area <= 0:
            raise ValueError(
                f"Segment {j + 1}: the channel geometry produces an "
                "invalid or zero-area boundary. Reduce gap/bend, or "
                "move the attachment points further apart."
            )

        segments.append(segment_polygon)

    if len(segments) != config.segments:
        raise ValueError(
            f"Requested {config.segments} segments, but "
            f"{len(segments)} were constructed for the current "
            "parameters. Please check the channel configuration."
        )

    for index, segment in enumerate(segments, start=1):
        validate_segment_polygon(segment, index)

    return segments, channel_curves, channel_slots


# ============================================================
# COORDINATE EXPORT
# ============================================================

def _export_turning_angles(xy: np.ndarray) -> np.ndarray:
    """Turning angle at every vertex of a closed ring."""
    xy = np.asarray(xy, dtype=float)
    prev = xy - np.roll(xy, 1, axis=0)
    nxt = np.roll(xy, -1, axis=0) - xy
    a1 = np.arctan2(prev[:, 1], prev[:, 0])
    a2 = np.arctan2(nxt[:, 1], nxt[:, 0])
    return np.abs((a2 - a1 + np.pi) % (2.0 * np.pi) - np.pi)


def _export_dp_keep(P: np.ndarray, tol: float, max_span: float) -> np.ndarray:
    """Douglas-Peucker reduction with a maximum exported point spacing."""
    P = np.asarray(P, dtype=float)
    n = len(P)
    if n <= 2:
        return np.ones(n, dtype=bool)

    keep = np.zeros(n, dtype=bool)
    keep[0] = keep[-1] = True
    s = np.concatenate([[0.0], np.cumsum(np.linalg.norm(np.diff(P, axis=0), axis=1))])
    stack = [(0, n - 1)]

    while stack:
        i, j = stack.pop()
        if j <= i + 1:
            continue

        a = P[i]
        seg = P[j] - P[i]
        L2 = float(seg @ seg)
        Q = P[i + 1:j] - a

        if L2 < 1e-24:
            d = np.linalg.norm(Q, axis=1)
        else:
            t = np.clip((Q @ seg) / L2, 0.0, 1.0)
            d = np.linalg.norm(Q - t[:, None] * seg, axis=1)

        k = int(np.argmax(d))
        if d[k] > tol:
            sp = i + 1 + k
        elif math.sqrt(L2) > max_span or (s[j] - s[i]) > 1.5 * max_span:
            sp = int(np.searchsorted(s, 0.5 * (s[i] + s[j])))
            sp = min(max(sp, i + 1), j - 1)
        else:
            continue

        keep[sp] = True
        stack += [(i, sp), (sp, j)]

    return keep


def _export_dedupe(xy: np.ndarray, flags_list: List[np.ndarray], threshold: float):
    """Remove consecutive near-duplicate vertices while preserving flagged corners."""
    if len(xy) <= 1:
        return xy, flags_list

    corner = flags_list[0]
    alive = np.ones(len(xy), dtype=bool)
    last = 0

    for i in range(1, len(xy)):
        if np.linalg.norm(xy[i] - xy[last]) < threshold:
            if corner[i] and not corner[last]:
                alive[last] = False
                last = i
            else:
                alive[i] = False
                for f in flags_list:
                    f[last] |= f[i]
        else:
            last = i

    if len(xy) > 1 and last != 0 and np.linalg.norm(xy[last] - xy[0]) < threshold:
        if corner[0] and not corner[last]:
            alive[last] = False
        else:
            alive[0] = False

    return xy[alive], [f[alive] for f in flags_list]


def _smooth_resample_branch(
    P: np.ndarray,
    ds_min: float,
    ds_max: float,
    theta_deg: float,
    grad: float,
) -> np.ndarray:
    """
    Resample one smooth branch by arc length.

    Local spacing = min(ds_max, radius * theta) (never below ds_min), then
    limited so it changes gradually along the branch. Points are placed so
    that spacing follows this field exactly (equal "spacing units" between
    points). The two end points are kept exactly. The result has no
    micro-edges and no abrupt spacing jumps, so a spline through it stays
    smooth.
    """
    P = np.asarray(P, dtype=float)
    d = np.linalg.norm(np.diff(P, axis=0), axis=1)
    P = P[np.concatenate([[True], d > 1e-12])]
    if len(P) < 2:
        return P

    s = np.concatenate([[0.0], np.cumsum(np.linalg.norm(np.diff(P, axis=0), axis=1))])
    L = float(s[-1])
    if L <= 1.5 * ds_min or len(P) < 3:
        return P[[0, -1]]

    # Fine uniform copy of the branch for stable curvature estimates.
    h = min(0.02, L / 60.0)
    nu = int(math.ceil(L / h)) + 1
    su = np.linspace(0.0, L, nu)
    U = np.column_stack([np.interp(su, s, P[:, 0]), np.interp(su, s, P[:, 1])])
    hh = su[1] - su[0]

    ang = np.unwrap(np.arctan2(np.gradient(U[:, 1]), np.gradient(U[:, 0])))
    k = max(1, int(round(0.25 / hh)))
    idx = np.arange(nu)
    lo = np.clip(idx - k, 0, nu - 1)
    hi = np.clip(idx + k, 0, nu - 1)
    span = np.maximum(su[hi] - su[lo], 1e-12)
    kappa = np.abs(ang[hi] - ang[lo]) / span
    radius = 1.0 / np.maximum(kappa, 1e-9)

    ds = np.clip(radius * math.radians(theta_deg), ds_min, ds_max)

    # Gradation: neighbouring spacing may change only gradually.
    for i in range(1, nu):
        ds[i] = min(ds[i], ds[i - 1] + grad * hh)
    for i in range(nu - 2, -1, -1):
        ds[i] = min(ds[i], ds[i + 1] + grad * hh)

    inv = 1.0 / ds
    ns = np.concatenate([[0.0], np.cumsum(0.5 * (inv[1:] + inv[:-1]) * hh)])
    total = float(ns[-1])
    m = max(1, int(round(total)))
    targets = np.linspace(0.0, total, m + 1)
    s_pts = np.interp(targets, ns, su)

    out = np.column_stack([np.interp(s_pts, su, U[:, 0]), np.interp(s_pts, su, U[:, 1])])
    out[0] = P[0]
    out[-1] = P[-1]
    return out


def reduce_export_ring(
    coordinates: np.ndarray,
    chord_mm: float,
    airfoil_boundary: LineString,
    channel_boundaries: List[LineString],
    tolerance_mm: float = EXPORT_POINT_TOLERANCE_MM,
    corner_angle_deg: float = EXPORT_CORNER_ANGLE_DEG,
    max_span_pct: float = EXPORT_MAX_SPAN_PCT,
    surface_tolerance_mm: float = EXPORT_SURFACE_TOLERANCE_MM,
):
    """
    Convert a dense closed segment perimeter into the exported point set:
      1. identify airfoil/channel junctions and hard turning corners,
      2. split the ring at those corners,
      3. reduce every smooth branch with Douglas-Peucker
         (airfoil-surface branches use surface_tolerance_mm, 0 = keep all;
          channel walls use tolerance_mm),
      4. enforce a maximum point spacing,
      5. preserve the corners and return explicit branch metadata.
    """
    xy = np.asarray(coordinates, dtype=float)
    xy = xy[np.all(np.isfinite(xy), axis=1)]
    if len(xy) < 4:
        raise ValueError("Segment has too few coordinates for point reduction.")

    if np.linalg.norm(xy[0] - xy[-1]) <= 1e-9 * max(chord_mm, 1.0):
        xy = xy[:-1]

    if len(xy) < 3:
        raise ValueError("Segment ring has too few unique vertices for point reduction.")

    scale = max(chord_mm, 1.0)
    eps = 1e-6 * scale
    turn = _export_turning_angles(xy)

    on_air = np.array([airfoil_boundary.distance(Point(p)) <= eps for p in xy], dtype=bool)
    if channel_boundaries:
        on_channel = np.array([
            min(line.distance(Point(p)) for line in channel_boundaries) <= eps
            for p in xy
        ], dtype=bool)
    else:
        on_channel = np.zeros(len(xy), dtype=bool)

    junction = on_air & on_channel
    corner = junction | (turn > math.radians(corner_angle_deg))

    xy, (corner, junction, on_air) = _export_dedupe(
        xy,
        [corner, junction, on_air],
        max(1e-9 * scale, MICRO_MERGE_MM),
    )

    cidx = np.flatnonzero(corner)
    if len(cidx) < 2:
        far = int(np.argmax(np.linalg.norm(xy - xy[0], axis=1)))
        corner[[0, far]] = True
        cidx = np.flatnonzero(corner)

    start = int(cidx[np.argmin(xy[cidx, 0])])
    order = (np.arange(len(xy)) + start) % len(xy)
    xy = xy[order]
    corner = corner[order]
    junction = junction[order]
    on_air = on_air[order]

    cidx = np.flatnonzero(corner)
    n = len(xy)
    ext = np.vstack([xy, xy[:1]])
    bounds = list(cidx) + [n]

    max_span = max_span_pct / 100.0 * chord_mm
    if max_span <= 0:
        raise ValueError("Maximum export point spacing must be greater than zero.")

    points_out = []
    corner_out = []
    branches = []

    for bidx in range(len(cidx)):
        i0, i1 = bounds[bidx], bounds[bidx + 1]
        branch_dense = ext[i0:i1 + 1]

        mid_index = (i0 + i1) // 2
        probe = Point(ext[mid_index]) if i1 - i0 >= 2 else Point(0.5 * (ext[i0] + ext[i1]))
        is_surface = airfoil_boundary.distance(probe) <= eps
        kind = "surface" if is_surface else "wall"

        if EXPORT_SMOOTH:
            pts_b = _smooth_resample_branch(
                branch_dense,
                ds_min=EXPORT_MIN_EDGE_MM,
                ds_max=max_span,
                theta_deg=EXPORT_ANGLE_STEP_DEG,
                grad=EXPORT_SPACING_GRADATION,
            )
        elif is_surface and surface_tolerance_mm <= 0.0:
            pts_b = branch_dense
        else:
            tol = surface_tolerance_mm if is_surface else tolerance_mm
            kept = np.flatnonzero(_export_dp_keep(branch_dense, tol, max_span))
            if len(kept) < 2:
                kept = np.array([0, len(branch_dense) - 1], dtype=int)
            pts_b = branch_dense[kept]

        r0 = len(points_out)
        points_out.extend(pts_b[:-1])
        corner_out.extend([True] + [False] * max(0, len(pts_b) - 2))

        branches.append({
            "id": bidx + 1,
            "kind": kind,
            "i0": r0,
            "i1": r0 + len(pts_b) - 1,
        })

    reduced = np.asarray(points_out, dtype=float)
    corner_out = np.asarray(corner_out, dtype=bool)

    if len(reduced) < 3:
        raise ValueError("Point reduction produced fewer than 3 vertices.")

    return reduced, corner_out, branches


def write_reduced_segment_csv(path: Path, coordinates: np.ndarray, section_id: int):
    """Write the reduced segment ring using the five-column CSV format."""
    path.parent.mkdir(parents=True, exist_ok=True)
    coordinates = clean_coordinates(coordinates, decimals=ANSYS_DECIMALS, closed=True)

    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Section_ID", "Point_ID", "X_mm", "Y_mm", "Z_mm"])
        for point_id, (x, y) in enumerate(coordinates, start=1):
            writer.writerow([section_id, point_id, f"{x:.6f}", f"{y:.6f}", "0.000000"])


def write_ansys_branch_txt(
    path: Path,
    reduced_coordinates: np.ndarray,
    branches: list,
    z: float = 0.0,
    group_offset: int = 0,
):
    """Write the reduced ring as separate ANSYS groups for each smooth branch."""
    path.parent.mkdir(parents=True, exist_ok=True)
    pts_all = np.asarray(reduced_coordinates, dtype=float)
    n = len(pts_all)
    lines = []

    for b in branches:
        gid = group_offset + int(b["id"])
        idx = np.arange(b["i0"], b["i1"] + 1) % n
        branch_pts = pts_all[idx]
        for point_id, (x, y) in enumerate(branch_pts, start=1):
            lines.append(f"{gid:d}\t{point_id:d}\t{x:.6f}\t{y:.6f}\t{z:.6f}")

    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="ascii", newline="\n")


def write_combined_reduced_ansys_txt(path: Path, reduced_segment_data: list, z: float = 0.0):
    """Write every segment, with unique ANSYS group IDs for every smooth branch."""
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = []
    gid = 0

    for reduced_coordinates, branches in reduced_segment_data:
        pts_all = np.asarray(reduced_coordinates, dtype=float)
        n = len(pts_all)
        for b in branches:
            gid += 1
            idx = np.arange(b["i0"], b["i1"] + 1) % n
            branch_pts = pts_all[idx]
            for point_id, (x, y) in enumerate(branch_pts, start=1):
                lines.append(f"{gid:d}\t{point_id:d}\t{x:.6f}\t{y:.6f}\t{z:.6f}")

    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="ascii", newline="\n")


def clean_coordinates(
    coordinates: np.ndarray,
    decimals: int = ANSYS_DECIMALS,
    closed: bool = True,
) -> np.ndarray:

    coordinates = np.asarray(coordinates, dtype=float)
    coordinates = coordinates[np.all(np.isfinite(coordinates), axis=1)]

    if len(coordinates) < 2:
        raise ValueError("Not enough valid coordinates to export.")

    if closed:
        if np.linalg.norm(coordinates[0] - coordinates[-1]) <= 1e-7:
            coordinates = coordinates[:-1]

    cleaned = [coordinates[0]]
    for point in coordinates[1:]:
        if np.linalg.norm(point - cleaned[-1]) > 1e-10:
            cleaned.append(point)

    coordinates = np.asarray(cleaned, dtype=float)
    coordinates = np.round(coordinates, decimals=decimals)

    cleaned = [coordinates[0]]
    for point in coordinates[1:]:
        if not np.array_equal(point, cleaned[-1]):
            cleaned.append(point)

    return np.asarray(cleaned, dtype=float)


def write_segment_csv(path: Path, coordinates: np.ndarray, section_id: int):
    """One CSV file: Section_ID, Point_ID, X_mm, Y_mm, Z_mm."""

    path.parent.mkdir(parents=True, exist_ok=True)

    coordinates = clean_coordinates(coordinates, decimals=ANSYS_DECIMALS, closed=True)

    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Section_ID", "Point_ID", "X_mm", "Y_mm", "Z_mm"])
        for point_id, (x, y) in enumerate(coordinates, start=1):
            writer.writerow([section_id, point_id, f"{x:.3f}", f"{y:.3f}", "0.000"])


def write_ansys_txt(path: Path, coordinates: np.ndarray, group_id: int = 1, closed: bool = True):
    """ANSYS DesignModeler 5-column TXT (Group Point X Y Z)."""

    path.parent.mkdir(parents=True, exist_ok=True)

    coordinates = clean_coordinates(coordinates, decimals=ANSYS_DECIMALS, closed=closed)

    lines = []
    for point_id, (x, y) in enumerate(coordinates, start=1):
        lines.append(f"{group_id} {point_id} {x:.3f} {y:.3f} 0")

    if closed:
        lines.append(f"{group_id} 0 0 0 0")

    path.write_text("\n".join(lines) + "\n", encoding="ascii", newline="\n")


def write_combined_ansys_txt(path: Path, segment_coordinate_list: List[np.ndarray], closed: bool = True):
    """Write ALL segments into one ANSYS file, one Group number per segment."""

    path.parent.mkdir(parents=True, exist_ok=True)

    lines = []

    for group_id, coordinates in enumerate(segment_coordinate_list, start=1):
        coordinates = clean_coordinates(coordinates, decimals=ANSYS_DECIMALS, closed=closed)

        for point_id, (x, y) in enumerate(coordinates, start=1):
            lines.append(f"{group_id} {point_id} {x:.3f} {y:.3f} 0")

        if closed:
            lines.append(f"{group_id} 0 0 0 0")

    path.write_text("\n".join(lines) + "\n", encoding="ascii", newline="\n")


def validate_segment_polygon(polygon: Polygon, index: int):
    """Reject a segment before export if its boundary is not a clean simple ring."""

    ring = polygon.exterior

    if not ring.is_ring:
        raise ValueError(f"Segment {index}: boundary is not a closed ring.")

    if not polygon.exterior.is_simple:
        raise ValueError(
            f"Segment {index}: boundary self-intersects. "
            "Reduce gap/bend or move the attachment points."
        )

    coords = np.asarray(ring.coords)

    if len(coords) < 4:
        raise ValueError(
            f"Segment {index}: boundary has too few points to be a "
            "valid closed curve."
        )

    if not np.all(np.isfinite(coords)):
        raise ValueError(
            f"Segment {index}: boundary contains invalid (NaN/inf) coordinates."
        )

    diffs = np.linalg.norm(np.diff(coords[:-1], axis=0), axis=1)

    if np.any(diffs <= 1e-9):
        raise ValueError(
            f"Segment {index}: boundary contains duplicate consecutive points."
        )


def write_centerline_csv(path: Path, coordinates: np.ndarray, section_id: int):
    path.parent.mkdir(parents=True, exist_ok=True)

    coordinates = clean_coordinates(coordinates, decimals=ANSYS_DECIMALS, closed=False)

    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Section_ID", "Point_ID", "X_mm", "Y_mm", "Z_mm"])
        for point_id, (x, y) in enumerate(coordinates, start=1):
            writer.writerow([section_id, point_id, f"{x:.3f}", f"{y:.3f}", "0.000"])


def write_parameters_csv(path: Path, config: GeneratorConfig):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Parameter", "Value", "Unit"])
        writer.writerow(["Chord", f"{config.chord_mm:.3f}", "mm"])
        writer.writerow(["Number_of_segments", config.segments, "-"])

        for channel in config.channels:
            writer.writerow([f"{channel.name}_Lower_attachment", f"{channel.lower_xc_percent:.3f}", "%c"])
            writer.writerow([f"{channel.name}_Upper_attachment", f"{channel.upper_xc_percent:.3f}", "%c"])
            writer.writerow([f"{channel.name}_Gap", f"{channel.gap_mm:.3f}", "mm"])
            writer.writerow([f"{channel.name}_Bend", f"{channel.bend_mm:.3f}", "mm"])
            writer.writerow([f"{channel.name}_Curve_method", channel.curve_method, "-"])


# ============================================================
# OUTPUT FOLDER
# ============================================================

def create_case_folder(input_file: Path) -> Path:
    root = input_file.parent / "segmented_output"
    root.mkdir(parents=True, exist_ok=True)

    number = 1

    while True:
        case = root / f"case_{number:03d}"

        if not case.exists():
            case.mkdir()
            return case

        number += 1


def save_all_outputs(
    case_dir: Path,
    original_points: np.ndarray,
    segments: List[Polygon],
    channel_curves: List[np.ndarray],
    channel_slots: List[Polygon],
    config: GeneratorConfig,
):
    """Save all CSV / ANSYS / parameter / report outputs."""

    segment_csv_dir = case_dir / "SEGMENT_CSV"
    ansys_dir = case_dir / "ANSYS_TXT"
    channel_dir = case_dir / "CHANNELS"
    reference_dir = case_dir / "REFERENCE"
    parameter_dir = case_dir / "PARAMETERS"

    for folder in [segment_csv_dir, ansys_dir, channel_dir, reference_dir, parameter_dir]:
        folder.mkdir(parents=True, exist_ok=True)

    # Reference airfoil line built from the SAME surface vertices the
    # segments are made of, so surface vertices are recognised exactly.
    up_fun, lo_fun = build_surface_functions(original_points)
    low_ref = np.column_stack([lo_fun.x, lo_fun.y])
    up_ref = np.column_stack([up_fun.x, up_fun.y])
    airfoil_line = LineString(np.vstack([low_ref, up_ref[::-1], low_ref[:1]]))
    channel_boundaries = [slot.boundary for slot in channel_slots]

    reduced_segment_data = []
    summary_rows = []

    for index, polygon in enumerate(segments, start=1):
        dense_coordinates = np.asarray(polygon.exterior.coords, dtype=float)

        reduced_coordinates, reduced_corner, branches = reduce_export_ring(
            dense_coordinates,
            chord_mm=config.chord_mm,
            airfoil_boundary=airfoil_line,
            channel_boundaries=channel_boundaries,
            tolerance_mm=EXPORT_POINT_TOLERANCE_MM,
            corner_angle_deg=EXPORT_CORNER_ANGLE_DEG,
            max_span_pct=EXPORT_MAX_SPAN_PCT,
            surface_tolerance_mm=EXPORT_SURFACE_TOLERANCE_MM,
        )

        reduced_segment_data.append((reduced_coordinates, branches))
        summary_rows.append((index, len(dense_coordinates) - 1, len(reduced_coordinates), len(branches)))

        write_reduced_segment_csv(
            segment_csv_dir / f"segment_{index}.csv",
            reduced_coordinates,
            section_id=index,
        )

        write_ansys_branch_txt(
            ansys_dir / f"segment_{index}.txt",
            reduced_coordinates,
            branches,
            z=0.0,
            group_offset=0,
        )

    write_combined_reduced_ansys_txt(
        ansys_dir / "ALL_SEGMENTS_COMBINED.txt",
        reduced_segment_data,
        z=0.0,
    )

    for index, curve in enumerate(channel_curves, start=1):
        write_centerline_csv(
            channel_dir / f"channel_{index}_centerline.csv",
            curve,
            section_id=100 + index,
        )

        write_ansys_txt(
            channel_dir / f"channel_{index}_centerline.txt",
            curve,
            group_id=1,
            closed=False,
        )

    write_segment_csv(reference_dir / "original_airfoil.csv", original_points, section_id=0)
    write_ansys_txt(reference_dir / "original_airfoil.txt", original_points, group_id=1, closed=True)

    write_parameters_csv(parameter_dir / "parameters.csv", config)

    with (parameter_dir / "parameters.json").open("w", encoding="utf-8") as f:
        json.dump(
            {
                "chord_mm": config.chord_mm,
                "segments": config.segments,
                "channels": [asdict(channel) for channel in config.channels],
                "s_departure_deg": S_DEPARTURE_DEG,
                "input_rebuild": {
                    "densify_input": DENSIFY_INPUT,
                    "densify_points": DENSIFY_POINTS,
                    "te_sharp": TE_SHARP,
                    "te_blend_pct": TE_BLEND_PCT,
                },
                "export_point_reduction": {
                    "smooth": EXPORT_SMOOTH,
                    "min_edge_mm": EXPORT_MIN_EDGE_MM,
                    "angle_step_deg": EXPORT_ANGLE_STEP_DEG,
                    "junction_min_edge_mm": JUNCTION_MIN_EDGE_MM,
                    "wall_tolerance_mm": EXPORT_POINT_TOLERANCE_MM,
                    "surface_tolerance_mm": EXPORT_SURFACE_TOLERANCE_MM,
                    "corner_angle_deg": EXPORT_CORNER_ANGLE_DEG,
                    "max_span_pct": EXPORT_MAX_SPAN_PCT,
                    "coordinate_decimals": ANSYS_DECIMALS,
                    "ansys_mode": "branches",
                },
            },
            f,
            indent=4,
        )

    with (case_dir / "geometry_report.txt").open("w", encoding="utf-8") as f:
        f.write("PARAMETRIC SEGMENTED AIRFOIL REPORT\n")
        f.write("=" * 72 + "\n\n")
        f.write(f"Chord                : {config.chord_mm:.3f} mm\n")
        f.write(f"Number of segments   : {config.segments}\n")
        f.write("\nINPUT AIRFOIL REBUILD\n")
        f.write("-" * 72 + "\n")
        f.write(f"Densify input         : {DENSIFY_INPUT}\n")
        f.write(f"Points per surface    : {DENSIFY_POINTS}\n")
        f.write(f"Sharp trailing edge   : {TE_SHARP} (blend {TE_BLEND_PCT:.2f} %c)\n")
        f.write("\nEXPORT POINT EXTRACTION\n")
        f.write("-" * 72 + "\n")
        f.write(f"Smooth export         : {EXPORT_SMOOTH} (min edge {EXPORT_MIN_EDGE_MM} mm, step {EXPORT_ANGLE_STEP_DEG} deg)\n")
        f.write(f"Wall tolerance        : {EXPORT_POINT_TOLERANCE_MM:.6f} mm\n")
        f.write(f"Surface tolerance     : {EXPORT_SURFACE_TOLERANCE_MM:.6f} mm\n")
        f.write(f"Corner angle          : {EXPORT_CORNER_ANGLE_DEG:.3f} deg\n")
        f.write(f"Maximum spacing       : {EXPORT_MAX_SPAN_PCT:.3f} %c\n")
        f.write(f"Coordinate decimals   : {ANSYS_DECIMALS}\n")
        f.write("ANSYS layout          : branches\n\n")
        f.write("Segment point counts\n")
        f.write("-" * 72 + "\n")
        for sid, dense_n, reduced_n, branch_n in summary_rows:
            f.write(
                f"Segment {sid}: {dense_n} dense boundary points -> "
                f"{reduced_n} exported points, {branch_n} ANSYS branches\n"
            )
        f.write("\nSEGMENT CSV FILES\n")
        f.write("-" * 72 + "\n")
        for index in range(1, config.segments + 1):
            f.write(f"SEGMENT_CSV/segment_{index}.csv\n")
        f.write("\nANSYS DESIGNMODELER FILES\n")
        f.write("-" * 72 + "\n")
        for index in range(1, config.segments + 1):
            f.write(f"ANSYS_TXT/segment_{index}.txt\n")
        f.write("ANSYS_TXT/ALL_SEGMENTS_COMBINED.txt\n")
        f.write("\nANSYS SETTINGS\n")
        f.write("-" * 72 + "\n")
        f.write("Definition       : From Coordinates File\n")
        f.write("Coordinates Unit : Millimeter\n")
        f.write("Base Plane       : XYPlane\n")
        f.write("TXT format       : Group Point X Y Z\n")
        f.write(f"TXT decimals     : {ANSYS_DECIMALS}\n")


def save_preview_png(
    original_points: np.ndarray,
    segments: List[Polygon],
    channel_curves: List[np.ndarray],
    config: GeneratorConfig,
    path: Path,
):
    """Save the same preview used in Page 2 as a PNG."""

    fig = make_preview_figure(original_points, segments, channel_curves, config)
    fig.savefig(path, dpi=220, bbox_inches="tight")
    fig.clear()


# ============================================================
# PREVIEW FIGURE
# ============================================================

def make_preview_figure(
    original_points: np.ndarray,
    segments: List[Polygon],
    channels: List[np.ndarray],
    config: GeneratorConfig,
):
    """Visual preview with segments, channel centre-lines and attachment points."""

    fig = Figure(figsize=(10.5, 6.2), dpi=100)
    ax = fig.add_subplot(111)

    original = np.asarray(original_points)

    ax.plot(original[:, 0], original[:, 1], linestyle="--", linewidth=1.0, label="Airfoil")

    for index, polygon in enumerate(segments, start=1):
        xy = np.asarray(polygon.exterior.coords)

        ax.plot(xy[:, 0], xy[:, 1], linewidth=2.0, label=f"Segment {index}")

        centroid = polygon.centroid

        ax.annotate(
            f"SEG {index}",
            (centroid.x, centroid.y),
            ha="center",
            va="center",
            fontsize=9,
            fontweight="bold",
        )

    upper_fun, lower_fun = build_surface_functions(original_points)

    for index, (channel, spec) in enumerate(zip(channels, config.channels), start=1):

        ax.plot(
            channel[:, 0],
            channel[:, 1],
            linestyle=":",
            linewidth=1.2,
            label=f"{spec.name} centreline",
        )

        lower_x = config.chord_mm * spec.lower_xc_percent / 100.0
        upper_x = config.chord_mm * spec.upper_xc_percent / 100.0

        lower_y = float(lower_fun(lower_x))
        upper_y = float(upper_fun(upper_x))

        ax.plot(lower_x, lower_y, marker="o", markersize=5)
        ax.plot(upper_x, upper_y, marker="o", markersize=5)

        ax.annotate(
            f"C{index} L\n{spec.lower_xc_percent:.0f}%c",
            (lower_x, lower_y),
            xytext=(0, -18),
            textcoords="offset points",
            ha="center",
            fontsize=8,
        )

        ax.annotate(
            f"C{index} U\n{spec.upper_xc_percent:.0f}%c",
            (upper_x, upper_y),
            xytext=(0, 18),
            textcoords="offset points",
            ha="center",
            fontsize=8,
        )

    y0, y1 = ax.get_ylim()
    pad = 0.22 * (y1 - y0)
    ax.set_ylim(y0 - pad, y1 + pad)

    ax.axhline(0, linewidth=0.8, linestyle="-.")

    for pct in (0, 25, 50, 75, 100):
        x = config.chord_mm * pct / 100.0

        ax.axvline(x, linewidth=0.45, linestyle=":", alpha=0.35)

        ax.text(x, ax.get_ylim()[0], f"{pct}%c", ha="center", va="bottom", fontsize=7)

    ax.set_xlabel("X [mm] / X/c")
    ax.set_ylabel("Y [mm]")
    ax.set_title(f"Segmented Airfoil Preview  |  Chord = {config.chord_mm:.1f} mm", pad=14)
    ax.grid(True, alpha=0.22)
    ax.set_aspect("equal", adjustable="box")
    ax.legend(loc="best", fontsize=7)

    fig.tight_layout()

    return fig


# ============================================================
# GUI
# ============================================================

class SegmentedAirfoilApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Parametric Segmented Airfoil Generator V10")
        self.geometry("1220x820")
        self.minsize(1080, 720)

        self.file_var = tk.StringVar()
        self.chord_var = tk.StringVar(value="100")
        self.segment_count_var = tk.StringVar(value="2")

        # Channel 1 defaults
        self.c1_lower = tk.StringVar(value="40")
        self.c1_upper = tk.StringVar(value="60")
        self.c1_gap = tk.StringVar(value="5")
        self.c1_bend = tk.StringVar(value="0")
        self.c1_method = tk.StringVar(value="S-Bezier")

        # Channel 2 defaults
        self.c2_lower = tk.StringVar(value="65")
        self.c2_upper = tk.StringVar(value="85")
        self.c2_gap = tk.StringVar(value="5")
        self.c2_bend = tk.StringVar(value="0")
        self.c2_method = tk.StringVar(value="S-Bezier")

        self.status_var = tk.StringVar(
            value=(
                "Step 1: Select your Excel airfoil coordinate file "
                "and enter the segmentation parameters."
            )
        )

        self.output_dir: Path | None = None
        self.generated_figure = None

        self._make_style()
        self._build_ui()

    # --------------------------------------------------------
    # STYLE
    # --------------------------------------------------------

    def _make_style(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("SubTitle.TLabel", font=("Segoe UI", 10))
        style.configure("Section.TLabelframe", padding=10)
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)

    # --------------------------------------------------------
    # BUILD GUI
    # --------------------------------------------------------

    def _build_ui(self):

        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="PARAMETRIC SEGMENTED AIRFOIL GENERATOR",
            style="Title.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            main,
            text=(
                "Excel input  →  parameterized segmentation  "
                "→  preview  →  separate CSV + ANSYS TXT"
            ),
            style="SubTitle.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True)

        self.page1 = ttk.Frame(self.notebook, padding=12)
        self.page2 = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.page1, text="1. INPUT & PARAMETERIZATION")
        self.notebook.add(self.page2, text="2. PREVIEW & SAVE")

        self._build_page1()
        self._build_page2()

        ttk.Separator(main).pack(fill="x", pady=8)

        ttk.Label(main, textvariable=self.status_var, wraplength=1150).pack(anchor="w")

    # --------------------------------------------------------
    # PAGE 1
    # --------------------------------------------------------

    def _build_page1(self):

        # A. Airfoil input
        input_frame = ttk.LabelFrame(self.page1, text="A. AIRFOIL INPUT", padding=10)
        input_frame.pack(fill="x", pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Excel file:").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(input_frame, textvariable=self.file_var).grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_file).grid(row=0, column=2, padx=6, pady=5)

        ttk.Label(
            input_frame,
            text="Required columns: Section_ID | Point_ID | X_mm | Y_mm | Z_mm",
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=4)

        # B. Basic geometry
        geometry_frame = ttk.LabelFrame(self.page1, text="B. BASIC GEOMETRY PARAMETERS", padding=10)
        geometry_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(geometry_frame, text="Chord").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(geometry_frame, textvariable=self.chord_var, width=14).grid(
            row=0, column=1, sticky="w", padx=6, pady=5
        )
        ttk.Label(geometry_frame, text="mm").grid(row=0, column=2, sticky="w", padx=2, pady=5)
        ttk.Label(geometry_frame, text="Number of segments").grid(row=0, column=3, sticky="e", padx=6, pady=5)

        segment_combo = ttk.Combobox(
            geometry_frame,
            textvariable=self.segment_count_var,
            values=["2", "3"],
            state="readonly",
            width=10,
        )
        segment_combo.grid(row=0, column=4, sticky="w", padx=6, pady=5)
        segment_combo.bind("<<ComboboxSelected>>", self.on_segment_count_changed)

        ttk.Label(
            geometry_frame,
            text="2 segments = 1 channel   |   3 segments = 2 channels",
        ).grid(row=1, column=0, columnspan=5, sticky="w", padx=6, pady=4)

        # C. Explanation
        explain = ttk.LabelFrame(
            self.page1,
            text="C. PARAMETERIZATION — WHAT EACH VALUE MEANS",
            padding=10,
        )
        explain.pack(fill="x", pady=(0, 10))

        explanation_text = (
            "Lower attachment [%c] : point where the channel starts on the lower airfoil surface.\n"
            "Upper attachment [%c] : point where the channel ends on the upper airfoil surface.\n"
            "Gap [mm]              : total width of the removed slot/channel.\n"
            "Bend [mm]             : Bezier / B-spline = sideways displacement of the centreline (0 = straight).\n"
            "                        S-Bezier = handle length of the S (0 = automatic).\n"
            "Curve method          : Bezier (C-shape), cubic B-spline, or S-Bezier (S-shape).\n\n"
            "Example for a 100 mm chord:\n"
            "40 %c = X = 40 mm, 60 %c = X = 60 mm."
        )

        ttk.Label(explain, text=explanation_text, justify="left").pack(anchor="w")

        # D. Channel table
        table_frame = ttk.LabelFrame(
            self.page1,
            text="D. CHANNEL / SEGMENTATION PARAMETERS",
            padding=10,
        )
        table_frame.pack(fill="x", pady=(0, 10))

        for col, title in enumerate(["Parameter", "Channel 1", "Channel 2"]):
            ttk.Label(table_frame, text=title, font=("Segoe UI", 10, "bold")).grid(
                row=0, column=col, sticky="w", padx=6, pady=5
            )

        self.c2_widgets = []

        rows = [
            ("Lower attachment [%c]", self.c1_lower, self.c2_lower),
            ("Upper attachment [%c]", self.c1_upper, self.c2_upper),
            ("Gap [mm]", self.c1_gap, self.c2_gap),
            ("Bend [mm]", self.c1_bend, self.c2_bend),
        ]

        for row_index, (label, first_var, second_var) in enumerate(rows, start=1):

            ttk.Label(table_frame, text=label).grid(row=row_index, column=0, sticky="w", padx=6, pady=4)

            ttk.Entry(table_frame, textvariable=first_var, width=16).grid(
                row=row_index, column=1, sticky="w", padx=6, pady=4
            )

            widget = ttk.Entry(table_frame, textvariable=second_var, width=16)
            widget.grid(row=row_index, column=2, sticky="w", padx=6, pady=4)

            self.c2_widgets.append(widget)

        ttk.Label(table_frame, text="Curve method").grid(row=5, column=0, sticky="w", padx=6, pady=4)

        ttk.Combobox(
            table_frame,
            textvariable=self.c1_method,
            values=CURVE_METHODS,
            state="readonly",
            width=14,
        ).grid(row=5, column=1, sticky="w", padx=6, pady=4)

        self.c2_method_widget = ttk.Combobox(
            table_frame,
            textvariable=self.c2_method,
            values=CURVE_METHODS,
            state="readonly",
            width=14,
        )
        self.c2_method_widget.grid(row=5, column=2, sticky="w", padx=6, pady=4)

        self.c2_widgets.append(self.c2_method_widget)

        # Generate
        button_frame = ttk.Frame(self.page1)
        button_frame.pack(fill="x", pady=10)

        ttk.Button(
            button_frame,
            text="GENERATE → SHOW PREVIEW",
            style="Big.TButton",
            command=self.generate,
        ).pack(side="left")

        ttk.Button(button_frame, text="Reset Parameters", command=self.reset_parameters).pack(
            side="left", padx=10
        )

        self.on_segment_count_changed()

    # --------------------------------------------------------
    # PAGE 2
    # --------------------------------------------------------

    def _build_page2(self):

        ttk.Label(
            self.page2,
            text=(
                "STEP 2 — GEOMETRY PREVIEW\n"
                "The left side shows the generated geometry. "
                "The right side shows exactly how it is parameterized."
            ),
            justify="left",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        body = ttk.Frame(self.page2)
        body.pack(fill="both", expand=True)

        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        plot_frame = ttk.LabelFrame(body, text="Generated Airfoil", padding=6)
        plot_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.plot_container = ttk.Frame(plot_frame)
        self.plot_container.pack(fill="both", expand=True)

        summary_frame = ttk.LabelFrame(body, text="Parameter Summary", padding=10)
        summary_frame.grid(row=0, column=1, sticky="nsew")

        self.summary_text = tk.Text(
            summary_frame, wrap="word", width=38, height=25, font=("Consolas", 9)
        )
        self.summary_text.pack(fill="both", expand=True)
        self.summary_text.configure(state="disabled")

        output_frame = ttk.LabelFrame(self.page2, text="Output Files", padding=10)
        output_frame.pack(fill="x", pady=(10, 0))

        self.output_text = tk.Text(output_frame, height=7, wrap="none", font=("Consolas", 9))
        self.output_text.pack(fill="x")
        self.output_text.configure(state="disabled")

        buttons = ttk.Frame(self.page2)
        buttons.pack(fill="x", pady=(10, 0))

        ttk.Button(
            buttons,
            text="← Back to Parameters",
            command=lambda: self.notebook.select(self.page1),
        ).pack(side="left")

        ttk.Button(buttons, text="Open Output Folder", command=self.open_output_folder).pack(
            side="left", padx=8
        )

    # --------------------------------------------------------
    # EVENTS
    # --------------------------------------------------------

    def on_segment_count_changed(self, event=None):
        try:
            count = int(self.segment_count_var.get())
        except ValueError:
            count = 2

        state = "normal" if count == 3 else "disabled"

        for widget in self.c2_widgets:
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass

        # Combobox must be "readonly" (not "normal") when enabled.
        if count == 3:
            try:
                self.c2_method_widget.configure(state="readonly")
            except tk.TclError:
                pass

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select airfoil coordinate Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )

        if path:
            self.file_var.set(path)
            self.status_var.set("Selected input file. Set the parameters and click Generate.")

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    @staticmethod
    def number(variable: tk.StringVar, label: str) -> float:

        try:
            value = float(variable.get())
        except ValueError as exc:
            raise ValueError(f"{label} must be a numeric value.") from exc

        if not math.isfinite(value):
            raise ValueError(f"{label} must be finite.")

        return value

    def build_config(self) -> GeneratorConfig:

        chord = self.number(self.chord_var, "Chord")

        if chord <= 0:
            raise ValueError("Chord must be greater than 0 mm.")

        segments = int(self.segment_count_var.get())

        channels = []

        first = ChannelSpec(
            name="Channel 1",
            lower_xc_percent=self.number(self.c1_lower, "Channel 1 lower attachment"),
            upper_xc_percent=self.number(self.c1_upper, "Channel 1 upper attachment"),
            gap_mm=self.number(self.c1_gap, "Channel 1 gap"),
            bend_mm=self.number(self.c1_bend, "Channel 1 bend"),
            curve_method=self.c1_method.get(),
        )

        channels.append(first)

        if segments == 3:
            second = ChannelSpec(
                name="Channel 2",
                lower_xc_percent=self.number(self.c2_lower, "Channel 2 lower attachment"),
                upper_xc_percent=self.number(self.c2_upper, "Channel 2 upper attachment"),
                gap_mm=self.number(self.c2_gap, "Channel 2 gap"),
                bend_mm=self.number(self.c2_bend, "Channel 2 bend"),
                curve_method=self.c2_method.get(),
            )

            channels.append(second)

        for channel in channels:

            if not (0.0 <= channel.lower_xc_percent <= 100.0):
                raise ValueError(f"{channel.name}: lower attachment must be 0..100 %c.")

            if not (0.0 <= channel.upper_xc_percent <= 100.0):
                raise ValueError(f"{channel.name}: upper attachment must be 0..100 %c.")

            if channel.lower_xc_percent >= channel.upper_xc_percent:
                raise ValueError(
                    f"{channel.name}: lower attachment must be "
                    "before upper attachment in X/c."
                )

            if channel.gap_mm <= 0:
                raise ValueError(f"{channel.name}: gap must be > 0 mm.")

        if segments == 3:
            if channels[0].upper_xc_percent >= channels[1].lower_xc_percent:
                raise ValueError(
                    "For 3 segments, Channel 1 must be located before "
                    "Channel 2 in the chordwise direction.\n\n"
                    "Example: C1 = 40–60 %c, C2 = 65–85 %c."
                )

        return GeneratorConfig(chord_mm=chord, segments=segments, channels=channels)

    # --------------------------------------------------------
    # GENERATION
    # --------------------------------------------------------

    def generate(self):

        if not self.file_var.get().strip():
            messagebox.showerror("Input required", "Select the Excel airfoil coordinate file.")
            return

        try:
            self.status_var.set("Reading airfoil coordinates...")
            self.update_idletasks()

            input_file = Path(self.file_var.get().strip())

            config = self.build_config()

            raw = read_airfoil_excel(input_file)

            scaled = scale_to_chord(raw, config.chord_mm)

            if DENSIFY_INPUT:
                self.status_var.set("Rebuilding smooth airfoil through the input points...")
                self.update_idletasks()
                scaled = densify_airfoil(scaled)

            self.status_var.set("Generating parameterized channels...")
            self.update_idletasks()

            segments, channel_curves, channel_slots = generate_geometry(scaled, config)

            self.status_var.set(
                "Reducing export points and saving segment CSV + ANSYS branch files..."
            )
            self.update_idletasks()

            case_dir = create_case_folder(input_file)

            save_all_outputs(case_dir, scaled, segments, channel_curves, channel_slots, config)

            save_preview_png(
                scaled,
                segments,
                channel_curves,
                config,
                case_dir / "segmented_airfoil_preview.png",
            )

            self.output_dir = case_dir

            self.show_preview(scaled, segments, channel_curves, config)

            self.notebook.select(self.page2)

            self.status_var.set(
                f"SUCCESS — generated {len(segments)} segments. "
                f"Files saved in: {case_dir}"
            )

        except Exception as exc:

            self.status_var.set(f"ERROR: {exc}")

            messagebox.showerror("Generation Error", str(exc))

    # --------------------------------------------------------
    # PREVIEW
    # --------------------------------------------------------

    def show_preview(
        self,
        original: np.ndarray,
        segments: List[Polygon],
        channels: List[np.ndarray],
        config: GeneratorConfig,
    ):

        for widget in self.plot_container.winfo_children():
            widget.destroy()

        fig = make_preview_figure(original, segments, channels, config)

        canvas = FigureCanvasTkAgg(fig, master=self.plot_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        self.generated_figure = fig

        lines = []

        lines.append("AIRFOIL\n-------")
        lines.append(f"Chord       : {config.chord_mm:.3f} mm")
        lines.append(f"Segments    : {config.segments}")
        lines.append(f"Smooth input: {DENSIFY_INPUT}")
        lines.append(f"Sharp TE    : {TE_SHARP if DENSIFY_INPUT else 'n/a'}")
        lines.append("")
        lines.append("CHANNEL PARAMETERS\n------------------")

        for channel in config.channels:

            x_lower = config.chord_mm * channel.lower_xc_percent / 100.0
            x_upper = config.chord_mm * channel.upper_xc_percent / 100.0

            lines.append(f"\n{channel.name}")
            lines.append(f"Lower : {channel.lower_xc_percent:.2f} %c (X={x_lower:.3f} mm)")
            lines.append(f"Upper : {channel.upper_xc_percent:.2f} %c (X={x_upper:.3f} mm)")
            lines.append(f"Gap   : {channel.gap_mm:.3f} mm")
            lines.append(f"Bend  : {channel.bend_mm:.3f} mm")
            lines.append(f"Curve : {channel.curve_method}")

        lines.append("")
        lines.append("SEGMENT FILES")
        lines.append("-------------")

        for index in range(1, config.segments + 1):
            lines.append(f"segment_{index}.csv")

        summary = "\n".join(lines)

        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", summary)
        self.summary_text.configure(state="disabled")

        output_lines = [
            "OUTPUT FOLDER:",
            str(self.output_dir),
            "",
            "SEGMENT CSV (exactly one file per segment):",
        ]

        for index in range(1, config.segments + 1):
            output_lines.append(f"  SEGMENT_CSV/segment_{index}.csv")

        output_lines.extend(["", "ANSYS TXT (same one-file-per-segment structure):"])

        for index in range(1, config.segments + 1):
            output_lines.append(f"  ANSYS_TXT/segment_{index}.txt")

        output_lines.extend(
            [
                "",
                "CHANNEL DATA:",
                "  CHANNELS/",
                "",
                "PARAMETERS:",
                "  PARAMETERS/parameters.csv",
                "  PARAMETERS/parameters.json",
                "",
                "PLOT:",
                "  segmented_airfoil_preview.png",
            ]
        )

        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", "\n".join(output_lines))
        self.output_text.configure(state="disabled")

    # --------------------------------------------------------
    # UTILITIES
    # --------------------------------------------------------

    def open_output_folder(self):

        if self.output_dir is None:
            messagebox.showinfo("No output", "Generate a segmented airfoil first.")
            return

        try:
            if os.name == "nt":
                os.startfile(str(self.output_dir))
            elif os.name == "posix":
                subprocess.Popen(["xdg-open", str(self.output_dir)])

        except Exception as exc:
            messagebox.showerror("Could not open folder", str(exc))

    def reset_parameters(self):

        self.chord_var.set("100")
        self.segment_count_var.set("2")

        self.c1_lower.set("40")
        self.c1_upper.set("60")
        self.c1_gap.set("5")
        self.c1_bend.set("0")
        self.c1_method.set("S-Bezier")

        self.c2_lower.set("65")
        self.c2_upper.set("85")
        self.c2_gap.set("5")
        self.c2_bend.set("0")
        self.c2_method.set("S-Bezier")

        self.on_segment_count_changed()

        self.status_var.set("Parameters reset. Select the Excel file and generate.")


def main():
    app = SegmentedAirfoilApp()
    app.mainloop()


if __name__ == "__main__":
    main()
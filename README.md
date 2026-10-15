# Parametric Segmented Airfoil Generator

A desktop tool (Python + Tkinter) that takes an airfoil coordinate file from Excel and cuts it into **2 or 3 segments** separated by constant-width, curved channels (slots). The result is exported as clean CSV files and as **ANSYS DesignModeler** coordinate files, ready to import in millimetres.

```
Excel airfoil  ->  parameterised channels  ->  preview  ->  segment CSV + ANSYS TXT
```

## Features

- Reads airfoil coordinates from an Excel sheet (`.xlsx` / `.xlsm`).
- Scales the airfoil to any chord length.
- Optional smooth, dense rebuild of the input airfoil (a spline passing exactly through every input point, cosine spaced) to remove long straight facets from sparse data.
- Optional sharp trailing edge (or keep the original blunt one).
- 1 or 2 channels, each defined by lower and upper attachment points (% chord), gap width, bend and curve type:
  - **Bezier**: C-shaped channel
  - **B-spline**: cubic B-spline centreline
  - **S-Bezier**: S-shaped channel that leaves and arrives at the surface at a set angle to the local tangent
- Channel walls are offset from the centreline by half the gap and cut exactly where they cross the airfoil surface.
- Smooth export: each boundary branch is resampled by arc length with gradually varying spacing (fine where the curve bends, coarse where it is flat), so splines through the exported points do not wiggle.
- No micro-edges: vertices closer than a minimum distance to a channel junction are dropped, so DesignModeler does not choke on sub-micron edges.
- Live preview and a saved PNG.
- Every case is saved in its own numbered output folder.

## Requirements

- Python 3.9 or newer
- Tkinter (included with most Python installers; on Debian/Ubuntu: `sudo apt install python3-tk`)
- Python packages:

```bash
pip install numpy scipy matplotlib openpyxl shapely
```

## Usage

```bash
python segmented_airfoil.py
```

(Use the actual file name of the script in this repository.)

1. **Browse** to your Excel airfoil file.
2. Enter the **chord** in mm and choose **2 or 3 segments** (2 segments = 1 channel, 3 segments = 2 channels).
3. Fill in the channel parameters.
4. Click **GENERATE -> SHOW PREVIEW**.
5. Check the preview on page 2 and use **Open Output Folder** to get the files.

### Input file format

The Excel file needs these columns (header row, case-insensitive):

| Section_ID | Point_ID | X_mm | Y_mm | Z_mm |
|-----------|----------|------|------|------|

- The sheet `Airfoil_Coordinates` is used if present, otherwise the first sheet.
- Only `X_mm` and `Y_mm` are used (2-D geometry).
- The contour should be a standard closed airfoil ordering (trailing edge -> leading edge -> trailing edge). At least 20 valid points are required.
- Blunt trailing edges (two points at max X) are supported.

### Channel parameters

| Parameter | Meaning |
|-----------|---------|
| Lower attachment [%c] | Where the channel starts on the lower surface |
| Upper attachment [%c] | Where the channel ends on the upper surface |
| Gap [mm] | Total width of the removed slot |
| Bend [mm] | Bezier / B-spline: sideways displacement of the centreline (0 = straight). S-Bezier: handle length of the S (0 = automatic) |
| Curve method | Bezier, B-spline or S-Bezier |

Rules:

- Lower attachment must be before upper attachment for each channel.
- With 3 segments, Channel 1 must lie entirely ahead of Channel 2 (e.g. C1 = 40-60 %c, C2 = 65-85 %c).
- Attachments must be further aft than the previous channel's on both surfaces.

For a 100 mm chord, 40 %c means X = 40 mm.

## Output

Each run creates `segmented_output/case_001`, `case_002`, ... next to the input Excel file:

```
segmented_output/case_001/
├── SEGMENT_CSV/
│   ├── segment_1.csv
│   ├── segment_2.csv
│   └── segment_3.csv          (3-segment runs only)
├── ANSYS_TXT/
│   ├── segment_1.txt ...
│   └── ALL_SEGMENTS_COMBINED.txt
├── CHANNELS/                  (centrelines as CSV and TXT)
├── REFERENCE/                 (original airfoil as CSV and TXT)
├── PARAMETERS/
│   ├── parameters.csv
│   └── parameters.json
├── geometry_report.txt
└── segmented_airfoil_preview.png
```

- **CSV** columns: `Section_ID, Point_ID, X_mm, Y_mm, Z_mm`
- **ANSYS TXT** columns: `Group Point X Y Z` (tab separated). Every smooth branch (airfoil surface piece or channel wall) is its own group, so DesignModeler creates one clean curve per branch.

## Importing into ANSYS DesignModeler

1. Set the unit system to **millimetre** before importing.
2. Use **Concept -> 3D Curve** (or **Create -> Point / Sketch from coordinates**) with *Definition: From Coordinates File*.
3. Select the TXT file from `ANSYS_TXT/`.
4. Base plane: **XYPlane**, coordinate unit: **Millimeter**.
5. Surface-body creation from the closed curves follows your normal workflow.

Importing in mm is intentional: the minimum edge length is kept above the DesignModeler tolerance at this scale.

## Configuration

Tuning constants are at the top of the script.

| Constant | Default | Purpose |
|----------|---------|---------|
| `DENSIFY_INPUT` | `True` | Rebuild the airfoil as a smooth, dense contour |
| `DENSIFY_POINTS` | `600` | Points per surface for the rebuild |
| `TE_SHARP` | `False` | Fade the thickness to a sharp trailing edge |
| `TE_BLEND_PCT` | `8.0` | % of chord over which the TE is faded |
| `EXPORT_SMOOTH` | `True` | Arc-length resampling with smooth spacing |
| `EXPORT_MIN_EDGE_MM` | `0.05` | Shortest exported edge |
| `EXPORT_ANGLE_STEP_DEG` | `1.5` | Maximum turning angle between exported points |
| `EXPORT_SPACING_GRADATION` | `0.25` | How fast spacing may change along a branch |
| `JUNCTION_MIN_EDGE_MM` | `0.05` | Drop vertices closer than this to a junction |
| `EXPORT_MAX_SPAN_PCT` | `2.0` | Maximum point spacing (% chord) |
| `S_DEPARTURE_DEG` | `18.0` | S-Bezier departure angle vs. local surface tangent |
| `S_AUTO_HANDLE_FRACTION` | `0.40` | Automatic S-Bezier handle length (fraction of P0->P3) |
| `ANSYS_DECIMALS` | `6` | Decimals in exported coordinates |

Set `EXPORT_SMOOTH = False` to use the older Douglas-Peucker point reduction (`EXPORT_POINT_TOLERANCE_MM`, `EXPORT_SURFACE_TOLERANCE_MM`, `EXPORT_CORNER_ANGLE_DEG`).

## Troubleshooting

| Message | Likely fix |
|---------|-----------|
| "Could not find where the channel wall meets the airfoil surface" | Reduce the gap, or move the attachment away from the leading/trailing edge |
| "The two channel walls overlap on the airfoil surface" | Reduce the gap or the bend |
| "...does not run from the lower to the upper surface" | Increase the separation between lower and upper attachment, or reduce the gap |
| "boundary self-intersects" / "invalid or zero-area boundary" | Reduce gap/bend or move attachment points further apart |
| "missing: ..." columns | Check the Excel header names |

## How it works (short)

1. Read and scale the airfoil, optionally rebuild it densely.
2. Split it into upper and lower surface functions `y(x)` (evaluated exactly on the polyline).
3. Build each channel centreline (Bezier, B-spline or S-Bezier) between the two attachment points.
4. Offset the centreline by +/- gap/2 to get two walls and intersect them with the surfaces.
5. Assemble each segment as a closed ring: surface -> wall -> surface -> wall, then validate it with Shapely.
6. Split each ring into smooth branches at junctions and corners, resample, and write the files.

## Version history

- **V10**: removed micro-edges for ANSYS mm import; smooth arc-length export resampling.
- **V9**: ANSYS mm import fixes.
- **V8**: exact surface evaluation; surface pieces keep their vertices.

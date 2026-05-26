# Fatbox — Notes and explanations

## What is Fatbox?

Fatbox is a Python toolbox for semi-automated fault extraction and structural analysis from various datasets including DEMs (Digital Elevation Models).

The project is structured as:
- `modules/` — 6 Python scripts containing all the functions
- `tutorials/` — Jupyter notebooks illustrating the workflow on 3 use cases
- `TDM1_DEM_30m_UTM37.tif` — the TanDEM-X DEM of the Afar region (your input file)

### The 6 modules

| Module | Role |
|---|---|
| `preprocessing.py` | Load and prepare the DEM (geotiff extraction, thresholding, skeletonisation) |
| `edits.py` | Extract and edit the fault network graph |
| `metrics.py` | Compute metrics (length, etc.) |
| `plots.py` | Visualise the network and results |
| `utils.py` | Low-level helper functions |
| `structural_analysis.py` | Measure geometric properties (strike, throw, dip, extension, displacement) |

---

## The DEM file (Afar region)

**File:** `TDM1_DEM_30m_UTM37.tif`
**Projection:** UTM Zone 37 (EPSG:32637)
**Resolution:** ~30.2 m/pixel
**Size:** 7343 × 11073 pixels (~222 km E–W × 334 km N–S)
**NW corner (UTM):** X = 716 922 m, Y = 1 440 548 m

Pixel-to-UTM conversion:
```
UTM_Easting  = 716922 + col * 30.21
UTM_Northing = 1440548 - row * 30.21
```

Inverse (to find pixel indices from UTM coordinates):
```
col = (UTM_Easting  - 716922) / 30.21
row = (1440548 - UTM_Northing) / 30.21
```

---

## The workflow — step by step

### 1. Load the DEM

```python
img_dem, lat, lon = preprocessing.extract_geotiff_and_coord(
    DEM_path,
    coord_north=0, coord_south=1000,
    coord_west=0,  coord_east=1000,
    system='grid coordinates'   # use 'geographic' for lat/lon
)
```

- Returns three NumPy arrays: elevation, latitudes, longitudes.
- `system='grid coordinates'` → pass pixel row/column indices (correct for UTM files).
- `system='geographic'` → pass decimal degrees lat/lon.
- You can extract a sub-tile to avoid processing the full ~7000×11000 array at once.

---

### 2. Fault extraction (image processing)

Four successive steps that convert the DEM into a set of 1-pixel-wide fault lines:

```
DEM → Gaussian smoothing → Canny edge detection → Clean small objects → Skeletonise
```

```python
smoothed = filters.gaussian(img_dem, sigma=s_smoothed)
edges    = feature.canny(smoothed, sigma=s_edges)
cleaned  = morphology.remove_small_objects(edges, connectivity=con_cleaned, min_size=msize_cleaned)
skeleton = skeletonize(cleaned)
```

#### Extraction parameters

**`s_smoothed`** — Gaussian smoothing sigma (pre-filter)
- Applied before edge detection to suppress pixel-level noise.
- Unit: pixels. `0` = no smoothing.
- Higher → smoother input, fewer spurious edges, but blurs real scarps.
- Typical range: 1–3.
- Afar: volcanic/lava surfaces are rough → try `2` if too many edges appear everywhere.

**`s_edges`** — Canny edge detection sigma
- Controls the scale at which the algorithm detects topographic discontinuities.
- Unit: pixels.
- Lower → detects fine, small features. Higher → only detects strong, broad scarps.
- Typical range: 2–6.
- Afar: if rivers and lava flow margins are detected as faults, increase this.

**`con_cleaned`** — connectivity for noise removal
- Defines what counts as "connected" when removing small objects.
- `2` = 8-connectivity (includes diagonals). Leave at `2`, do not change.

**`msize_cleaned`** — minimum object size
- Any connected group of edge-pixels smaller than this is deleted.
- Unit: pixels. `30 px × 30 m/px = 900 m minimum fault length`.
- Higher → cleaner result but loses short faults.
- Afar: real normal faults are typically several km long → can go up to 50–100.

#### Tuning guide

| Problem visible in the plots | Fix |
|---|---|
| Too many tiny edges everywhere | Increase `s_smoothed` and/or `msize_cleaned` |
| Real fault scarps not detected | Decrease `s_edges` or `s_smoothed` |
| Rivers/lava flows detected as faults | Increase `s_edges` |
| Short noise segments survive cleaning | Increase `msize_cleaned` |
| Long faults broken into many pieces | Decrease `msize_cleaned`, or use `connect_close_components` in filtering |

---

### 3. Classification — build the fault network graph

Each pixel of the skeleton becomes a graph node. Nodes belonging to the same connected component are linked by edges. Each component = one fault.

```python
G = nx.Graph()
# ... build nodes and edges from markers array ...
G = edits.label_components(G)
G = edits.remove_self_edge(G)
```

Key parameter: `distance_threshold = 1.5` (pixels) — nodes closer than this are connected. Leave at 1.5.

---

### 4. Filter the network

Chain of edits to clean the raw graph:

| Function | What it does |
|---|---|
| `connect_close_components(G, 4.0)` | Merges fault endpoints that are close but not yet connected |
| `simplify(G, 3)` | Reduces node density (keeps 1 node every 3) |
| `split_triple_junctions(G, 6)` | Separates forked fault tips |
| `loop_cut_U_global(G, img_dem)` | Removes U-shaped artefacts using the DEM to decide which side is the scarp |
| `remove_small_components(G, 5)` | Discards components with fewer than 5 nodes |
| `label_components(G)` | Relabels components to have consecutive numbers (run after any edit) |

---

### 5. Border processing

Removes fault segments whose cross-section would fall outside the DEM array (avoids index errors during structural analysis).

```python
J = structural_analysis.strike_edges(J)
J = metrics.compute_edge_length_meters(J, resolution)
J = metrics.compute_edge_length(J)
J = structural_analysis.filter_out_edges(J, img_dem, d=12)
J = edits.remove_small_components(J, 5)
J = edits.label_components(J)
```

Parameter `d` = half-length of the cross-section in pixels.
`d = 12 px × 30 m = 360 m on each side.`

---

### 6. Structural analysis

For each fault segment a topographic cross-section is drawn perpendicular to the fault, extending `d` pixels on each side of the segment midpoint. From the elevation profile the code computes:

- **Strike** — azimuth direction of the fault (degrees)
- **Throw** — maximum vertical offset across the scarp (metres)
- **Natural dip** — apparent dip angle measured from the DEM
- **Extension** — horizontal component of displacement (uses constrained dip if `use_natural_dip=False`)
- **Displacement** — total slip vector magnitude (√(throw² + extension²))
- **Dip direction** — which side is the hanging wall (east or west)

Key settings:
```python
d               = 12   # cross-section half-length in pixels
dip_constrain   = 60   # degrees, used when use_natural_dip=False
use_natural_dip = False
```

`use_natural_dip=False` is recommended when fault scarps are eroded (measured dip would be too low and overestimate extension).

**`summary_matrix`** — one row per fault component, 7 columns:

| Col | Content | Unit |
|---|---|---|
| 0 | Component label | — |
| 1 | Mean strike | degrees |
| 2 | Total length | metres (or pixels) |
| 3 | Maximum natural dip | degrees |
| 4 | Mean throw | metres |
| 5 | Mean extension | metres |
| 6 | Maximum displacement | metres |

---

### 7. Output plots

| Plot | What it shows |
|---|---|
| Rose diagram | Distribution of fault strike directions |
| Length histogram | Distribution of fault lengths |
| N≥length | Number of faults longer than a given length (power-law check) |
| Extension/throw/displacement map | Edge attribute coloured on top of the DEM |
| D/L scatter | Max displacement vs length for all faults |
| D/L profile | Displacement along a single fault (bell-shape = isolated fault) |

---

## Tuto_DEM_GLO30_hand_map_notebook — what it does

This tutorial shows how to **import a manually-drawn fault map** into the Fatbox workflow instead of using automated edge detection.

The key difference is the **mapping step**: instead of running Gaussian + Canny on the DEM, you load a GeoTIFF raster of your hand-drawn faults (black-and-white, fault pixels = 1, background = 0) and convert it to a graph with `cv2.connectedComponents`.

After that, the rest of the pipeline is identical: filtering → border processing → structural analysis → plots.

It also overlays both networks (hand-mapped vs automated) on the same hillshade to compare them visually.

### How to create the hand-map GeoTIFF in QGIS

1. Draw fault traces as a vector layer (multiline shapefile).
2. Rasterize with GDAL "Rasterize (vector to raster)":
   - Burn value: `1`
   - Resolution: same as DEM (30 m)
   - Extent: same bounding box as your DEM tile
   - Nodata value: `0`
   - Pre-initialize with: `0`
3. Export as GeoTIFF (same CRS as DEM).
4. Load in Fatbox with `preprocessing.extract_geotiff_and_coord`.

---

## Using hand mapping to tune extraction parameters

A hand-mapped fault layer gives ground truth to compare against the automated extraction. Recommended workflow:

1. Map a small representative tile in QGIS (10–20 km², a few hours of work).
2. Export as GeoTIFF at 30 m resolution with the same pixel extent as your DEM tile.
3. Load both the automated result and the hand map into `Tuto_DEM_GLO30_hand_map_notebook`.
4. Overlay the two networks — adjust `s_smoothed`, `s_edges`, `msize_cleaned` until they match.
5. Once satisfying, apply the tuned parameters to the full DEM.

**Best test tile:** choose an area with both clear fault scarps (easy cases) and noisy terrain (lava flows, wadis) so that the tuned parameters generalise to the rest of the DEM.

---

## Adapted notebook for the Afar DEM

**File:** `tutorials/digital_elevation_models/Tuto_DEM_Afar_notebook.ipynb`

Changes from the original tutorial:
- No Google Colab cells (no Drive mount, no pip install)
- `path_folder` → `/home/guiguizz/stage/modules`
- `save_path` → `/home/guiguizz/stage/tutorials/digital_elevation_models`
- DEM loaded directly from `TDM1_DEM_30m_UTM37.tif` (not from pre-saved `.npy` files)
- Coordinate system: `'grid coordinates'` (pixel indices, correct for UTM projection)
- Default tile: first 1000×1000 pixels (≈ 30 km × 30 km NW corner)
- Output filenames use `*_afar*` suffix

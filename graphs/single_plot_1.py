import pandas as pd
import numpy as np

# =============================
# (1) PNG: Matplotlib
# =============================
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize

# =============================
# (2) HTML: Plotly
# =============================
import plotly.graph_objects as go

# =============================
# НАСТРОЙКИ ФАЙЛА EXCEL
# =============================
file  = r"D:\10_results\02_Scopus\2.1.xlsx"
sheet = "SWEEP_2"

out_png  = r"C:\Users\Balt_\Desktop\surface_one.png"
out_html = r"C:\Users\Balt_\Desktop\surface_one.html"

# НОМЕР ГРАФИКА (1..4) -> выбор блока Z
PLOT_NO = 1  # <-- меняйте: 1,2,3,4

# ВЫБОР НАБОРА ДАННЫХ (1/2/3)
DATASET = 1  # <-- меняйте: 1, 2 или 3

# ПОДПИСИ ОСЕЙ / УГЛЫ
X_LABEL = "Максимальный ток разряда<br>Maximum discharge current"
Y_LABEL = "Емкость СНЭ, %<br>ESS capacity, %"
Z_LABEL = "Расходы, млн. руб.<br>Costs, million RUB"

ELEV, AZIM = 22, 35  # Matplotlib
CAMERA_EYE = dict(x=0.941, y=2.075, z=0.861)
CAMERA_UP = dict(x=0, y=0, z=1)
CAMERA_CENTER = dict(x=0, y=0, z=0)

X_ROW = 2
TRI_LIMIT = 1.0000001

if DATASET == 1:
    Y_START_ROW, Y_END_ROW = 3, 9
    X_START_COL, X_END_COL = 2, 12
    Z_BLOCKS = [(3, 9), (14, 20), (25, 31), (36, 42)]
    SHAPE_MODE = "rect"
elif DATASET == 2:
    Y_START_ROW, Y_END_ROW = 3, 12
    X_START_COL, X_END_COL = 2, 12
    Z_BLOCKS = [(3, 12), (17, 26), (31, 40), (45, 54)]
    SHAPE_MODE = "rect"
elif DATASET == 3:
    Y_START_ROW, Y_END_ROW = 3, 13
    X_START_COL, X_END_COL = 2, 12
    Z_BLOCKS = [(3, 13), (18, 28), (33, 43), (48, 58)]
    SHAPE_MODE = "triangle"
else:
    raise ValueError("DATASET должен быть 1, 2 или 3")

# =============================
# ВСПОМОГАТЕЛЬНОЕ: robust float
# =============================
def to_float(v):
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "###":
            return np.nan
        s = s.replace(" ", "").replace("\u00A0", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return np.nan
    return np.nan

def series_to_float(arr):
    s = pd.Series(arr).astype(str).str.strip()
    s = s.replace({"None": "", "nan": "", "NaN": "", "###": ""})
    s = s.str.replace("\u00A0", "", regex=False).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").to_numpy(dtype=float)

def build_mesh_from_grid(x_vals, y_vals, Zgrid):
    """
    Делает треугольную сетку (i,j,k) из регулярной таблицы Z (с NaN-дырками),
    как в вашем втором коде (Mesh3d + Scatter3d).
    """
    Xg, Yg = np.meshgrid(x_vals, y_vals)

    valid = np.isfinite(Zgrid)
    rr, cc = np.where(valid)

    vx = Xg[rr, cc].astype(float)
    vy = Yg[rr, cc].astype(float)
    vz = Zgrid[rr, cc].astype(float)

    idx = -np.ones(Zgrid.shape, dtype=int)
    idx[rr, cc] = np.arange(vx.size)

    I, J, K = [], [], []
    rows, cols = Zgrid.shape

    for r in range(rows - 1):
        for c in range(cols - 1):
            a = idx[r, c]
            b = idx[r, c + 1]
            d = idx[r + 1, c]
            e = idx[r + 1, c + 1]

            present = [t for t in (a, b, e, d) if t != -1]

            if len(present) == 4:
                I += [a, a]
                J += [b, e]
                K += [e, d]
            elif len(present) == 3:
                I.append(present[0])
                J.append(present[1])
                K.append(present[2])

    return vx, vy, vz, I, J, K

# =============================
# ПРОВЕРКИ / ВЫБОР Z
# =============================
if not (1 <= PLOT_NO <= len(Z_BLOCKS)):
    raise ValueError(f"PLOT_NO должен быть 1..{len(Z_BLOCKS)}, получено: {PLOT_NO}")

Z_START_ROW, Z_END_ROW = Z_BLOCKS[PLOT_NO - 1]

n_y = Y_END_ROW - Y_START_ROW + 1
if (Z_END_ROW - Z_START_ROW + 1) != n_y:
    raise RuntimeError(
        f"Высота блока Z ({Z_START_ROW}:{Z_END_ROW}) должна совпадать с числом строк Y ({Y_START_ROW}:{Y_END_ROW})."
    )

# =============================
# ЧТЕНИЕ ДАННЫХ (по Excel-координатам row/col)
# =============================
df = pd.read_excel(file, sheet_name=sheet, header=None)

x_raw = df.iloc[X_ROW - 1, X_START_COL - 1:X_END_COL].to_numpy()
x = series_to_float(x_raw)

y_raw = df.iloc[Y_START_ROW - 1:Y_END_ROW, 0].to_numpy()
y = series_to_float(y_raw)

z_raw = df.iloc[Z_START_ROW - 1:Z_END_ROW, X_START_COL - 1:X_END_COL].to_numpy()
z = np.vectorize(to_float, otypes=[float])(z_raw).astype(float)

if np.isnan(x).any():
    bad = np.where(np.isnan(x))[0].tolist()
    raise ValueError(f"X содержит NaN (нечисловые/пустые значения). Индексы: {bad}")
if np.isnan(y).any():
    bad = np.where(np.isnan(y))[0].tolist()
    raise ValueError(f"Y содержит NaN (нечисловые/пустые значения). Индексы: {bad}")

X, Y = np.meshgrid(x, y)

Z = z.copy()
if SHAPE_MODE == "triangle":
    mask = np.isnan(Z) | ((X + Y) > TRI_LIMIT)
    Z[mask] = np.nan

# =============================
# ЦВЕТА (общие)
# =============================
mpl_colors = ["limegreen", "yellowgreen", "yellow", "orange", "red", "maroon"]
cmap = LinearSegmentedColormap.from_list("custom_scale", mpl_colors, N=256)

plotly_colorscale = [
    [0.00, "limegreen"],
    [0.20, "yellowgreen"],
    [0.40, "yellow"],
    [0.60, "orange"],
    [0.80, "red"],
    [1.00, "maroon"],
]

# =============================
# TICKS (X через одну; Y — корректно для дробных/целых)
# =============================
x_tick_pos = x[::2]
x_tick_lbl = [f"{v:g}" for v in x_tick_pos]

y_tick_pos = y[::2]
y_tick_lbl = []
for v in y_tick_pos:
    vf = float(v)
    if np.isfinite(vf) and abs(vf - round(vf)) < 1e-9:
        y_tick_lbl.append(str(int(round(vf))))
    else:
        y_tick_lbl.append(f"{vf:.2f}".rstrip("0").rstrip(".").replace(".", ","))

# =====================================================================
# 1) HTML (Plotly) — Mesh3d + points (как во втором коде)
# =====================================================================
zmin = float(np.nanmin(Z))
zmax = float(np.nanmax(Z))

vx, vy, vz, I, J, K = build_mesh_from_grid(x, y, Z)

fig_html = go.Figure()

mesh = go.Mesh3d(
    x=vx, y=vy, z=vz,
    i=I, j=J, k=K,
    intensity=vz,
    colorscale=plotly_colorscale,
    cmin=zmin, cmax=zmax,
    showscale=False,
    flatshading=False,
    hovertemplate="x=%{x}<br>y=%{y}<br>z=%{z:.2f}<extra></extra>",
    showlegend=False,
)

points = go.Scatter3d(
    x=vx, y=vy, z=vz,
    mode="markers",
    marker=dict(
        size=3,
        opacity=1.0,
        color=vz,
        colorscale=plotly_colorscale,
        cmin=zmin,
        cmax=zmax,
        showscale=False
    ),
    hovertemplate="x=%{x}<br>y=%{y}<br>z=%{z:.2f}<extra></extra>",
    showlegend=False
)

fig_html.add_trace(mesh)
fig_html.add_trace(points)

fig_html.update_layout(
    template="plotly_white",
    width=1400,
    height=900,
    margin=dict(l=60, r=120, t=60, b=60),
    scene=dict(
        xaxis=dict(
            title=dict(text=X_LABEL, font=dict(size=14)),
            tickmode="array",
            tickvals=x_tick_pos.tolist(),
            ticktext=x_tick_lbl,
            tickfont=dict(size=12),
        ),
        yaxis=dict(
            title=dict(text=Y_LABEL, font=dict(size=14)),
            tickmode="array",
            tickvals=y_tick_pos.tolist(),
            ticktext=y_tick_lbl,
            tickfont=dict(size=12),
        ),
        zaxis=dict(
            title=dict(text=Z_LABEL, font=dict(size=14)),
            tickfont=dict(size=12),
        ),
        camera=dict(
            eye=CAMERA_EYE,
            up=CAMERA_UP,
            center=CAMERA_CENTER
        ),
        aspectmode="manual",
        aspectratio=dict(x=1.2, y=1.2, z=0.7),
    ),
)

# =============================
# DEBUG: показывать текущую камеру в HTML при вращении
# =============================
post_script = r"""
(function () {
  function pickGraphDiv() {
    // иногда Plotly создаёт несколько div; берём первый с графиком
    var els = document.querySelectorAll('.plotly-graph-div');
    if (!els || !els.length) return null;
    return els[0];
  }

  function getCam(gd) {
    try {
      // 1) самое надёжное — _fullLayout.scene.camera
      if (gd && gd._fullLayout && gd._fullLayout.scene && gd._fullLayout.scene.camera)
        return gd._fullLayout.scene.camera;
    } catch(e) {}

    try {
      // 2) layout.scene.camera
      if (gd && gd.layout && gd.layout.scene && gd.layout.scene.camera)
        return gd.layout.scene.camera;
    } catch(e) {}

    return null;
  }

  var gd = pickGraphDiv();
  if (!gd) return;

  var box = document.createElement('pre');
  box.id = 'camBox';
  box.style.position = 'fixed';
  box.style.right = '10px';
  box.style.bottom = '10px';
  box.style.maxWidth = '45vw';
  box.style.maxHeight = '45vh';
  box.style.overflow = 'auto';
  box.style.padding = '8px 10px';
  box.style.margin = '0';
  box.style.background = 'rgba(255,255,255,0.90)';
  box.style.border = '1px solid #999';
  box.style.borderRadius = '6px';
  box.style.fontSize = '12px';
  box.style.fontFamily = 'Consolas, Menlo, monospace';
  box.style.zIndex = 9999;
  document.body.appendChild(box);

  function round3(x){ return (typeof x === 'number') ? Math.round(x*1000)/1000 : x; }

  function render(cam){
    if (!cam) { box.textContent = 'camera: (нет данных)'; return; }
    var eye = cam.eye || {};
    var up = cam.up || {};
    var center = cam.center || {};
    box.textContent =
`CAMERA_EYE = dict(x=${round3(eye.x)}, y=${round3(eye.y)}, z=${round3(eye.z)})
CAMERA_UP  = dict(x=${round3(up.x)},  y=${round3(up.y)},  z=${round3(up.z)})
CAMERA_CENTER = dict(x=${round3(center.x)}, y=${round3(center.y)}, z=${round3(center.z)})`;
  }

  // старт
  render(getCam(gd));

  // апдейт при любых relayout (rotate/zoom/pan)
  gd.on('plotly_relayout', function(){
    render(getCam(gd));
  });
})();
"""

fig_html.write_html(out_html, include_plotlyjs="cdn", auto_open=True, post_script=post_script)
print(f"✅ HTML сохранён и открыт: {out_html}")

# =====================================================================
# 2) PNG (Matplotlib) — один график
# =====================================================================
mm_to_in = 1 / 25.4
fig_width_in = 360 * mm_to_in
fig_height_in = 220 * mm_to_in

fig = plt.figure(figsize=(fig_width_in, fig_height_in), dpi=300)
ax = fig.add_axes([0.08, 0.10, 0.72, 0.80], projection="3d")

vmin = np.nanmin(Z)
vmax = np.nanmax(Z)
norm = Normalize(vmin=vmin, vmax=vmax)

Zmask = np.ma.masked_invalid(Z)
Z_for_colors = np.where(np.isfinite(Z), Z, vmin)

ax.plot_surface(
    X, Y, Zmask,
    facecolors=cmap(norm(Z_for_colors)),
    rstride=1, cstride=1,
    linewidth=0.3,
    edgecolor="black",
    antialiased=True
)

ax.set_xlabel("Максимальный ток разряда\nMaximum discharge current", labelpad=12, fontsize=12)
ax.set_ylabel("Емкость СНЭ, %\nESS capacity, %", labelpad=12, fontsize=12)
ax.set_zlabel("Расходы, млн. руб.\nCosts, million RUB", labelpad=30, fontsize=12, rotation=180)

ax.tick_params(axis="x", pad=4, labelsize=10)
ax.tick_params(axis="y", pad=4, labelsize=10)
ax.tick_params(axis="z", pad=6, labelsize=10)

ax.set_xticks(x_tick_pos)
ax.set_xticklabels(x_tick_lbl, fontsize=10)

ax.set_yticks(y_tick_pos)
ax.set_yticklabels(y_tick_lbl, fontsize=10)

ax.view_init(elev=ELEV, azim=AZIM)

fig.patch.set_facecolor("white")
ax.set_facecolor("white")
plt.subplots_adjust(left=0.05, right=0.92, top=0.95, bottom=0.08)

plt.savefig(out_png, dpi=300)
plt.close(fig)

print(f"✅ PNG сохранён: {out_png}")
print(f"DATASET={DATASET} | PLOT_NO={PLOT_NO} | SHAPE_MODE={SHAPE_MODE} | Z rows {Z_START_ROW}:{Z_END_ROW}")

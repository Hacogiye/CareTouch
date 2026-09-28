# -*- coding: utf-8 -*-
"""CareTouch brand assets: gradient backgrounds, logo, charts.
Palette: Healing Teal + Warm Touch (orange)."""
import os, math
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

OUT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(OUT, "assets")
os.makedirs(ASSETS, exist_ok=True)

# ---------- palette ----------
DARK   = (11, 59, 52)     # 0B3B34 deep pine
DARK2  = (14, 74, 65)     # 0E4A41
PRIMARY= (14, 124, 107)   # 0E7C6B
TEALMID= (42, 168, 143)   # 2AA88F
TEALSOFT=(158, 217, 204)  # 9ED9CC
TINT   = (232, 245, 241)  # E8F5F1
ACCENT = (232, 89, 12)    # E8590C
ACCENTL= (247, 150, 90)   # lighter coral
TEXT   = (22, 48, 43)     # 16302B
MUTED  = (92, 115, 112)   # 5C7370

F = r"C:\Windows\Fonts"
def font(name, size):
    return ImageFont.truetype(os.path.join(F, name), size)

# ---------- helpers ----------
def diag_gradient(w, h, c1, c2):
    """Smooth diagonal gradient c1(top-left) -> c2(bottom-right)."""
    x = np.linspace(0, 1, w)[None, :]
    y = np.linspace(0, 1, h)[:, None]
    t = np.clip((x + y) / 2, 0, 1) ** 1.1
    arr = np.zeros((h, w, 3), dtype=np.float64)
    for i in range(3):
        arr[..., i] = c1[i] + (c2[i] - c1[i]) * t
    return Image.fromarray(arr.astype(np.uint8), "RGB")

def ripple(draw, cx, cy, r0, r1, steps, color, base_img, alpha0=90):
    """Concentric fading rings (touch ripple)."""
    for i in range(steps, 0, -1):
        r = r0 + (r1 - r0) * i / steps
        a = int(alpha0 * (1 - i / steps) ** 1.6)
        if a <= 2: continue
        layer = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
        ld = ImageDraw.Draw(layer)
        ld.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color + (a,), width=max(2, int(r1/steps/2.2)))
        base_img.alpha_composite(layer)

def soft_circle(base, cx, cy, r, color, alpha):
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color + (alpha,))
    layer = layer.filter(ImageFilter.GaussianBlur(r * 0.25))
    base.alpha_composite(layer)

def pair_of_circles(base, cx, cy, r, gap, c_a, c_b, alpha=200):
    """Brand motif: two overlapping translucent circles = Care + Touch."""
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.ellipse([cx - gap/2 - 2*r, cy - r, cx - gap/2, cy + r], fill=c_a + (alpha,))
    d.ellipse([cx + gap/2, cy - r, cx + gap/2 + 2*r, cy + r], fill=c_b + (alpha,))
    base.alpha_composite(layer)

W, H = 2667, 1500  # 13.33 x 7.5 in @200dpi

# ================= COVER BG =================
img = diag_gradient(W, H, DARK2, DARK).convert("RGBA")
# giant ghost ripples
soft_circle(img, int(W*0.86), int(H*0.30), 560, TEALMID, 26)
soft_circle(img, int(W*0.90), int(H*0.72), 420, PRIMARY, 30)
soft_circle(img, int(W*0.10), int(H*0.88), 380, DARK2, 60)
ripple(ImageDraw.Draw(img), int(W*0.86), int(H*0.30), 300, 900, 7, TEALSOFT, img, alpha0=60)
# brand pair motif on right
pair_of_circles(img, int(W*0.855), int(H*0.42), 200, 10, PRIMARY, ACCENT, alpha=110)
pair_of_circles(img, int(W*0.855), int(H*0.42), 120, 8, TEALSOFT, ACCENTL, alpha=90)
img.convert("RGB").save(os.path.join(ASSETS, "bg_cover.png"), quality=92)

# ================= CLOSING BG =================
img = diag_gradient(W, H, DARK, DARK2).convert("RGBA")
soft_circle(img, int(W*0.13), int(H*0.10), 430, PRIMARY, 26)
soft_circle(img, int(W*0.88), int(H*0.85), 460, DARK2, 70)
ripple(ImageDraw.Draw(img), int(W*0.13), int(H*0.10), 240, 760, 7, TEALSOFT, img, alpha0=45)
pair_of_circles(img, int(W*0.13), int(H*0.10), 130, 8, TEALSOFT, ACCENT, alpha=80)
img.convert("RGB").save(os.path.join(ASSETS, "bg_close.png"), quality=92)

# ================= LIGHT BG with soft blobs (for section dividers) =================
img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
soft_circle(img, int(W*0.92), int(H*0.08), 420, TINT, 255)
soft_circle(img, int(W*0.06), int(H*0.95), 380, TINT, 235)
pair_of_circles(img, int(W*0.93), int(H*0.10), 120, 7, TEALSOFT, ACCENTL, alpha=70)
img.convert("RGB").save(os.path.join(ASSETS, "bg_light.png"), quality=92)

# ================= LOGO PNG (for DOCX title) =================
def make_logo(path, scale=4, on_dark=False):
    lw, lh = 1450 * scale // 4, 360 * scale // 4
    img = Image.new("RGBA", (lw, lh), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # two overlapping circles
    r = int(lh * 0.30)
    cy = int(lh * 0.42)
    ox = int(lw * 0.055)
    gap = int(r * 0.9)
    circles = Image.new("RGBA", img.size, (0,0,0,0))
    cd = ImageDraw.Draw(circles)
    cd.ellipse([ox - gap, cy - r, ox + gap, cy + r], fill=(PRIMARY if not on_dark else TEALSOFT) + (255,))
    cd.ellipse([ox + 4, cy - r, ox + 4 + 2*r + gap - 4, cy + r], fill=ACCENT + (235,))
    img.alpha_composite(circles)
    tx = ox + 2*r + gap + int(lw*0.045)
    f_big = font("segoeuib.ttf", int(lh*0.44))
    d.text((tx, int(lh*0.16)), "Care", font=f_big, fill=(DARK if not on_dark else (255,255,255)))
    w_care = d.textlength("Care", font=f_big)
    d.text((tx + w_care, int(lh*0.16)), "Touch", font=f_big, fill=ACCENT)
    # tagline
    f_tag = font("segoeui.ttf", int(lh*0.135))
    d.text((tx + 4, int(lh*0.68)), "Healing Touch, Healthy Life", font=f_tag,
           fill=(MUTED if not on_dark else TEALSOFT))
    img.save(path)
    return path

make_logo(os.path.join(ASSETS, "logo.png"))
make_logo(os.path.join(ASSETS, "logo_dark.png"), on_dark=True)
print("backgrounds + logos done")

# ================= CHARTS (matplotlib, brand style) =================
import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager
import matplotlib.pyplot as plt

for f in ["segoeui.ttf", "segoeuib.ttf", "segoeuisl.ttf", "segoeuil.ttf"]:
    p = os.path.join(F, f)
    if os.path.exists(p):
        font_manager.fontManager.addfont(p)
plt.rcParams.update({
    "font.family": "Segoe UI",
    "text.color": "#16302B", "axes.edgecolor": "#CBE2DB",
    "axes.labelcolor": "#5C7370", "xtick.color": "#5C7370",
    "ytick.color": "#5C7370", "axes.linewidth": 1.0,
    "figure.facecolor": "white", "axes.facecolor": "white",
    "svg.fonttype": "none",
})
DPI = 200

def style_ax(ax, ygrid=True):
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    if ygrid:
        ax.grid(axis="y", color="#E3EEEA", linewidth=1)
        ax.set_axisbelow(True)
    ax.tick_params(length=0)

# --- Chart 1: aging population ---
fig, ax = plt.subplots(figsize=(7.4, 3.6), dpi=DPI)
cats = ["2024", "2025", "2030 (dự báo)"]
vals = [14.2, 16.1, 18.0]
bars = ax.bar(cats, vals, width=0.52, color=["#9ED9CC", "#0E7C6B", "#E8590C"], zorder=3)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v + 0.35, f"{v:.1f}".replace(".", ",") + " tr.",
            ha="center", fontsize=13, fontweight="bold", color="#16302B")
ax.set_ylim(0, 21)
ax.set_ylabel("Triệu người (60+)", fontsize=11)
style_ax(ax)
fig.tight_layout(pad=0.6)
fig.savefig(os.path.join(ASSETS, "chart_aging.png"), bbox_inches="tight")
plt.close(fig)

# --- Chart 2: health spending ---
fig, ax = plt.subplots(figsize=(7.4, 3.6), dpi=DPI)
cats = ["2019", "2025", "2028 (dự báo)"]
vals = [17.4, 27.5, 34.1]
bars = ax.bar(cats, vals, width=0.52, color=["#9ED9CC", "#0E7C6B", "#E8590C"], zorder=3)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v + 0.8, f"{v:.1f}".replace(".", ","),
            ha="center", fontsize=13, fontweight="bold", color="#16302B")
ax.set_ylim(0, 40)
ax.set_ylabel("Tỷ USD", fontsize=11)
style_ax(ax)
fig.tight_layout(pad=0.6)
fig.savefig(os.path.join(ASSETS, "chart_spend.png"), bbox_inches="tight")
plt.close(fig)

# --- Chart 3: wellness growth comparison ---
fig, ax = plt.subplots(figsize=(7.4, 3.4), dpi=DPI)
cats = ["Việt Nam", "Ấn Độ", "Trung Quốc", "Indonesia"]
vals = [15.6, 8.1, 7.6, 7.5]
bars = ax.barh(cats[::-1], vals[::-1], height=0.55,
               color=["#9ED9CC", "#9ED9CC", "#9ED9CC", "#E8590C"], zorder=3)
for b, v in zip(bars, vals[::-1]):
    ax.text(v + 0.25, b.get_y() + b.get_height()/2,
            ("+" + f"{v:.1f}".replace(".", ",")) + "%/năm",
            va="center", fontsize=12, fontweight="bold", color="#16302B")
ax.set_xlim(0, 19)
style_ax(ax, ygrid=False)
ax.grid(axis="x", color="#E3EEEA", linewidth=1)
ax.set_axisbelow(True)
fig.tight_layout(pad=0.6)
fig.savefig(os.path.join(ASSETS, "chart_wellness.png"), bbox_inches="tight")
plt.close(fig)

# --- Chart 4: 12-month revenue projection ---
fig, ax = plt.subplots(figsize=(8.6, 4.0), dpi=DPI)
months = [f"T{i}" for i in range(1, 13)]
rev = [22, 30, 38, 44, 50, 56, 62, 68, 73, 78, 83, 88]
fixed = 48
colors = ["#E8590C" if i == 4 else ("#0E7C6B" if v >= fixed else "#9ED9CC") for i, v in enumerate(rev)]
bars = ax.bar(months, rev, width=0.62, color=colors, zorder=3)
ax.axhline(fixed, color="#16302B", linewidth=1.6, linestyle=(0, (5, 3)), zorder=4)
ax.text(11.45, fixed + 1.6, "Chi phí cố định 48 tr./tháng", ha="right", fontsize=11,
        color="#16302B", fontweight="bold")
for b, v in zip(bars, rev):
    ax.text(b.get_x() + b.get_width()/2, v + 1.6, str(v), ha="center",
            fontsize=10.5, color="#5C7370")
ax.annotate("Hòa vốn\n(tháng 5)", xy=(4, 50), xytext=(2.1, 74),
            fontsize=12, fontweight="bold", color="#E8590C", ha="center",
            arrowprops=dict(arrowstyle="->", color="#E8590C", lw=1.8))
ax.set_ylim(0, 100)
ax.set_ylabel("Triệu đồng/tháng", fontsize=11)
style_ax(ax)
fig.tight_layout(pad=0.6)
fig.savefig(os.path.join(ASSETS, "chart_revenue.png"), bbox_inches="tight")
plt.close(fig)

print("charts done:", os.listdir(ASSETS))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge
import numpy as np
from PIL import Image, ImageDraw, ImageFont

NAVY = "#1F3864"
TEAL = "#14B8A6"

def render_mark(path, bg=None, size_px=1200):
    fig = plt.figure(figsize=(size_px/300, size_px/300), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(-1.7, 1.7)
    ax.set_ylim(-1.7, 1.7)
    ax.set_aspect("equal")
    ax.axis("off")
    if bg:
        fig.patch.set_facecolor(bg)
        ax.set_facecolor(bg)
    else:
        fig.patch.set_alpha(0)
        ax.patch.set_alpha(0)

    # Navy ring — open "Q" loop, gap at bottom-right (from 280 to 350 deg)
    ring = Wedge((0, 0), 1.0, -10, 280, width=0.26, facecolor=NAVY, edgecolor=None, capstyle="round")
    ax.add_patch(ring)
    # round the ring's two open ends
    for ang in (-10, 280):
        rad = np.radians(ang)
        rr = 0.87
        ax.add_patch(plt.Circle((rr*np.cos(rad), rr*np.sin(rad)), 0.13, facecolor=NAVY))

    # Teal tail — Q's diagonal flick, radial from inner ring out through the gap
    ang = np.radians(-45)
    r0, r1 = 0.50, 1.62
    x0, y0 = r0*np.cos(ang), r0*np.sin(ang)
    x1, y1 = r1*np.cos(ang), r1*np.sin(ang)
    ax.plot([x0, x1], [y0, y1], color=TEAL, linewidth=34, solid_capstyle="round")

    fig.savefig(path, dpi=300, transparent=(bg is None))
    plt.close(fig)

render_mark("mark_transparent_raw.png", bg=None)
render_mark("mark_white_raw.png", bg="#FFFFFF")

def square_crop_resize(src, dst, size, bg_for_flatten=None):
    img = Image.open(src).convert("RGBA")
    bbox = img.getbbox()
    img = img.crop(bbox)
    # pad to square with margin
    w, h = img.size
    side = int(max(w, h) * 1.35)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0) if bg_for_flatten is None else bg_for_flatten)
    canvas.alpha_composite(img, ((side - w)//2, (side - h)//2))
    canvas = canvas.resize((size, size), Image.LANCZOS)
    canvas.save(dst)
    return canvas

square_crop_resize("mark_transparent_raw.png", "Qflow_Logo_Mark_300x300_transparent.png", 300, bg_for_flatten=None)
square_crop_resize("mark_white_raw.png", "Qflow_Logo_Mark_300x300_white.png", 300, bg_for_flatten=(255,255,255,255))
square_crop_resize("mark_transparent_raw.png", "Qflow_Logo_Mark_1000x1000_transparent.png", 1000, bg_for_flatten=None)

# Lockup: mark + wordmark
mark_big = Image.open("mark_transparent_raw.png").convert("RGBA")
bbox = mark_big.getbbox()
mark_big = mark_big.crop(bbox)
mw, mh = mark_big.size
target_mark_h = 900
scale = target_mark_h / mh
mark_big = mark_big.resize((int(mw*scale), int(mh*scale)), Image.LANCZOS)
mw, mh = mark_big.size

canvas_w, canvas_h = 2400, 1300
canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 255))
canvas.alpha_composite(mark_big, ((canvas_w - mw)//2, 40))

d = ImageDraw.Draw(canvas)
f_brand = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 150)
f_tag = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 58)
t1, t2 = "QFLOW CONSULTING", "ISO & QUALITY MANAGEMENT"
b1 = d.textbbox((0,0), t1, font=f_brand)
b2 = d.textbbox((0,0), t2, font=f_tag)
d.text(((canvas_w-(b1[2]-b1[0]))/2, mh+70), t1, font=f_brand, fill=NAVY)
d.text(((canvas_w-(b2[2]-b2[0]))/2, mh+70+(b1[3]-b1[1])+40), t2, font=f_tag, fill=TEAL)
canvas.save("Qflow_Logo_Lockup_white.png")

print("done")

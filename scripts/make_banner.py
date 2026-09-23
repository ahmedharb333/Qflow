from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math

NAVY = (31, 56, 100, 255)
NAVY_DARK = (18, 34, 64, 255)
TEAL = (20, 184, 166, 255)
WHITE = (255, 255, 255, 255)

SCALE = 4
W, H = 1128 * SCALE, 191 * SCALE

img = Image.new("RGBA", (W, H), NAVY)

# Diagonal gradient navy -> darker navy (left to right)
grad = Image.new("L", (W, 1))
for x in range(W):
    t = x / W
    grad.putpixel((x, 0), int(255 * t))
grad = grad.resize((W, H))
dark_layer = Image.new("RGBA", (W, H), NAVY_DARK)
img = Image.composite(dark_layer, img, grad.point(lambda v: int(v * 0.55)))

d = ImageDraw.Draw(img, "RGBA")

# Decorative teal ring / flow-arc motif on the right side (subtle)
def ring(cx, cy, r, width, color, start=0, end=360):
    d.arc([cx - r, cy - r, cx + r, cy + r], start=start, end=end, fill=color, width=width)

teal_soft = (20, 184, 166, 55)
teal_soft2 = (20, 184, 166, 30)
navy_soft = (255, 255, 255, 18)

ring(W * 1.00, H * 0.5, H * 0.95, int(9 * SCALE), teal_soft, 30, 300)
ring(W * 1.05, H * 0.5, H * 1.30, int(6 * SCALE), teal_soft2, 50, 260)
ring(W * 0.965, H * 0.5, H * 0.62, int(6 * SCALE), navy_soft, -30, 190)

# thin dotted checklist lines, far right, very subtle (quality/process motif)
for i in range(3):
    y = H * (0.30 + i * 0.20)
    x0 = W * 0.955
    for k in range(4):
        x = x0 + k * (16 * SCALE)
        if x > W - 12 * SCALE:
            break
        d.ellipse([x, y - 3*SCALE, x + 6*SCALE, y + 3*SCALE], fill=(20,184,166,70))

# NOTE: no logo mark drawn here — LinkedIn overlays the Page's own logo circle
# over the bottom-left corner of the banner automatically; keep that zone clear.

# Text block
d = ImageDraw.Draw(img, "RGBA")
f_title = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", int(H * 0.28))
f_sub = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", int(H * 0.115))

text_x = int(W * 0.235)  # clears the LinkedIn logo-overlap zone (bottom-left)
title = "QFLOW CONSULTING"
subtitle = "ISO & Quality Management Consulting  |  Audit-Ready. Always."

tb = d.textbbox((0, 0), title, font=f_title)
sb = d.textbbox((0, 0), subtitle, font=f_sub)
title_h = tb[3] - tb[1]
sub_h = sb[3] - sb[1]
gap = int(H * 0.06)
total_h = title_h + gap + sub_h
start_y = (H - total_h) // 2 - tb[1]

d.text((text_x, start_y), title, font=f_title, fill=WHITE)
d.text((text_x, start_y + title_h + gap - sb[1] + tb[1]), subtitle, font=f_sub, fill=(154, 217, 207, 255))

img = img.resize((1128, 191), Image.LANCZOS)
img.convert("RGB").save("Qflow_LinkedIn_Banner_1128x191.png", quality=95)
print("done", img.size)

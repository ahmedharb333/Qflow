from PIL import Image, ImageDraw, ImageFont
import math

NAVY = (31, 56, 100, 255)
TEAL = (20, 184, 166, 255)

SCALE = 4
SIZE = 500 * SCALE
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
cx, cy = SIZE / 2, SIZE / 2
r_outer = SIZE * 0.46
r_inner = SIZE * 0.36

d.ellipse([cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer], outline=NAVY, width=int(SIZE * 0.018))
d.ellipse([cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner], outline=TEAL, width=int(SIZE * 0.012))

font_big = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", int(SIZE * 0.075))
font_small = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", int(SIZE * 0.045))


def draw_arc_text(text, radius, start_deg, end_deg, font, fill, flip=False):
    n = len(text)
    if n == 0:
        return
    step = (end_deg - start_deg) / max(n - 1, 1)
    for i, ch in enumerate(text):
        ang = math.radians(start_deg + step * i)
        x = cx + radius * math.cos(ang)
        y = cy + radius * math.sin(ang)
        glyph = Image.new("RGBA", (int(SIZE * 0.16), int(SIZE * 0.16)), (0, 0, 0, 0))
        gd = ImageDraw.Draw(glyph)
        gd.text((glyph.width / 2, glyph.height / 2), ch, font=font, fill=fill, anchor="mm")
        rot_deg = start_deg + step * i + 90 + (180 if flip else 0)
        glyph = glyph.rotate(-rot_deg, resample=Image.BICUBIC, center=(glyph.width / 2, glyph.height / 2))
        img.alpha_composite(glyph, (int(x - glyph.width / 2), int(y - glyph.height / 2)))


draw_arc_text("QFLOW CONSULTING", r_outer * 0.86, 200, 340, font_big, NAVY)
draw_arc_text("ABU DHABI - U.A.E.", r_outer * 0.86, 160, 20, font_small, TEAL, flip=True)

# center mark
cf = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", int(SIZE * 0.11))
d.text((cx, cy), "ISO", font=cf, fill=NAVY, anchor="mm")

img.resize((500, 500), Image.LANCZOS).save("Qflow_Company_Stamp.png")
print("done")

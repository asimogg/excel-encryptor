"""Uygulama ikonunu (assets/icon.png) çizer. Sadece build sırasında gerekir: pip install pillow"""
import os
from PIL import Image, ImageDraw, ImageFont
S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# arka plan: dikey degrade (Excel yeşili -> koyu)
top, bot = (33, 115, 70), (13, 59, 40)
grad = Image.new("RGB", (1, S))
for y in range(S):
    t = y / S
    grad.putpixel((0, y), tuple(int(a + (b - a) * t) for a, b in zip(top, bot)))
grad = grad.resize((S, S))
mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.225), fill=255)
img.paste(grad, (0, 0), mask)

# kilit
cx, W = S // 2, int(S * 0.30)
body_t, body_b = int(S * 0.46), int(S * 0.76)
d.rounded_rectangle([cx - W, body_t, cx + W, body_b], radius=int(S * 0.055),
                    fill=(255, 255, 255, 255))
sh_w, sh_top = int(S * 0.185), int(S * 0.235)
d.arc([cx - sh_w, sh_top, cx + sh_w, sh_top + 2 * sh_w], 180, 360,
      fill=(255, 255, 255, 255), width=int(S * 0.062))
d.rectangle([cx - sh_w - int(S*0.031), sh_top + sh_w, cx - sh_w + int(S*0.031), body_t + 10],
            fill=(255, 255, 255, 255))
d.rectangle([cx + sh_w - int(S*0.031), sh_top + sh_w, cx + sh_w + int(S*0.031), body_t + 10],
            fill=(255, 255, 255, 255))
# anahtar deliği
kx, ky, kr = cx, int(S * 0.575), int(S * 0.042)
d.ellipse([kx - kr, ky - kr, kx + kr, ky + kr], fill=(21, 92, 57, 255))
d.polygon([(kx - kr * 0.55, ky), (kx + kr * 0.55, ky),
           (kx + kr * 0.32, ky + kr * 2.6), (kx - kr * 0.32, ky + kr * 2.6)],
          fill=(21, 92, 57, 255))

# XLSX etiketi
for path in ["/System/Library/Fonts/Supplemental/Arial Bold.ttf",
             "/System/Library/Fonts/Helvetica.ttc", "/Library/Fonts/Arial.ttf"]:
    try:
        font = ImageFont.truetype(path, int(S * 0.105)); break
    except OSError:
        font = None
if font:
    d.text((cx, int(S * 0.855)), "XLSX", font=font, fill=(255, 255, 255, 235), anchor="mm")

img.save(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png"))
print("ikon çizildi")

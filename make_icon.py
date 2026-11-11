# -*- coding: utf-8 -*-
"""Génère icon.ico + icon.png pour Trafik Saveurs (biscuiterie & confiserie)."""
import math
from PIL import Image, ImageDraw

S = 4  # supersampling
N = 512
W = H = N * S

def ss(): return Image.new("RGBA", (W, H), (0, 0, 0, 0))

BROWN   = (59, 27, 18, 255)
DARK    = (41, 19, 13, 255)
GOLD    = (233, 169, 0, 255)
GOLD_L  = (255, 197, 46, 255)
CREAM   = (246, 243, 236, 255)
RED     = (184, 50, 38, 255)
GREEN   = (33, 122, 75, 255)
WHITE   = (255, 255, 255, 255)

img = ss()
d = ImageDraw.Draw(img)

# --- Fond : cercle dégradé brun ---
cx = cy = W // 2
R = int(W * 0.475)
d.ellipse([cx-R, cy-R, cx+R, cy+R], fill=BROWN)
# anneau doré
ring = int(R * 0.055)
d.ellipse([cx-R, cy-R, cx+R, cy+R], outline=GOLD, width=ring)
r2 = R - int(R*0.115)
d.ellipse([cx-r2, cy-r2, cx+r2, cy+r2], outline=(120, 70, 40, 255), width=max(2, ring//3))
# halo supérieur
for i in range(40):
    a = int(10 * (1 - i/40))
    rr = r2 - i * (r2//90)
    d.ellipse([cx-rr, cy-rr-int(rr*0.35), cx+rr, cy+rr-int(rr*0.35)], fill=(90, 45, 28, a))

def rounded_rect(box, rad, fill=None, outline=None, width=1):
    d.rounded_rectangle(box, radius=rad, fill=fill, outline=outline, width=width)

# --- Cookie (biscuit) en haut à gauche ---
ck_r = int(R * 0.30)
ckx, cky = cx - int(R*0.42), cy - int(R*0.40)
# contour grignoté (morceau mangé)
d.ellipse([ckx-ck_r, cky-ck_r, ckx+ck_r, cky+ck_r], fill=GOLD_L)
d.ellipse([ckx-ck_r, cky-ck_r, ckx+ck_r, cky+ck_r], outline=BROWN, width=int(ck_r*0.10))
bite_r = int(ck_r*0.42)
d.ellipse([ckx+int(ck_r*0.55)-bite_r, cky-int(ck_r*0.75)-bite_r,
           ckx+int(ck_r*0.55)+bite_r, cky-int(ck_r*0.75)+bite_r], fill=DARK)
for (px, py, pr) in [(-0.35,-0.15,0.13),(0.1,-0.4,0.11),(-0.05,0.3,0.14),(0.35,0.15,0.10),(-0.45,0.25,0.09)]:
    chip = int(ck_r*pr)
    d.ellipse([ckx+int(ck_r*px)-chip, cky+int(ck_r*py)-chip,
               ckx+int(ck_r*px)+chip, cky+int(ck_r*py)+chip], fill=BROWN)

# --- Bonbon (candy) en haut à droite ---
cd_r = int(R * 0.20)
cdx, cdy = cx + int(R*0.45), cy - int(R*0.42)
d.ellipse([cdx-cd_r, cdy-cd_r, cdx+cd_r, cdy+cd_r], fill=RED, outline=BROWN, width=int(cd_r*0.12))
# rayures blanches courbes approx.
mask = Image.new("L", (W, H), 0)
md = ImageDraw.Draw(mask)
md.ellipse([cdx-cd_r, cdy-cd_r, cdx+cd_r, cdy+cd_r], fill=255)
stripes = Image.new("RGBA", (W, H), (0,0,0,0))
sd = ImageDraw.Draw(stripes)
sw = int(cd_r*0.34)
for off in (-int(cd_r*0.9), int(cd_r*0.25)):
    sd.line([(cdx+off-cd_r, cdy+cd_r*1.4), (cdx+off+cd_r, cdy-cd_r*1.4)], fill=WHITE, width=sw)
img.paste(Image.composite(stripes, img, mask), (0,0), mask)
# emballage (deux triangles)
tw = int(cd_r*0.75)
for sgn in (-1, 1):
    bx = cdx + sgn*(cd_r + int(cd_r*0.55))
    d.polygon([(bx, cdy), (bx - sgn*tw, cdy - tw), (bx - sgn*tw, cdy + tw)],
              fill=GOLD, outline=BROWN)

# --- Barre de chocolat en bas à droite ---
bw, bh = int(R*0.62), int(R*0.34)
bbx, bby = cx + int(R*0.30), cy + int(R*0.42)
d.rotate(-12) if False else None
bar = Image.new("RGBA", (bw*2, bh*2), (0,0,0,0))
bd = ImageDraw.Draw(bar)
bd.rounded_rectangle([0, 0, bw*2-1, bh*2-1], radius=int(bh*0.25), fill=(110, 55, 30, 255), outline=BROWN, width=int(bh*0.10))
for i in range(2):
    for j in range(3):
        x0 = int(bw*2*(i*0.5) + bw*0.10); y0 = int(bh*2*(j*(1/3)) + bh*0.12)
        x1 = int(bw*2*(i*0.5) + bw*0.90); y1 = int(bh*2*(j*(1/3)) + bh*0.88)
        bd.rounded_rectangle([x0, y0, x1, y1], radius=int(bh*0.12), fill=(140, 72, 40, 255), outline=BROWN, width=int(bh*0.05))
bar = bar.rotate(12, expand=True, resample=Image.BICUBIC)
img.alpha_composite(bar, (bbx - bar.width//2, bby - bar.height//2))

# --- Épi/blé doré stylisé ? -> Remplacé par panier/caddie simple au centre-bas gauche ---
# Caddie (panier) minimaliste
pcx, pcy = cx - int(R*0.42), cy + int(R*0.40)
pw, ph = int(R*0.42), int(R*0.30)
lw = max(6, int(R*0.045))
pd = d
pd.line([(pcx-pw//2, pcy-ph//3), (pcx+pw//2, pcy-ph//3)], fill=CREAM, width=lw)          # rebord
pd.line([(pcx-pw//2, pcy-ph//3), (pcx-pw//3, pcy+ph//2)], fill=CREAM, width=lw)         # côté gauche
pd.line([(pcx+pw//2, pcy-ph//3), (pcx+pw//3, pcy+ph//2)], fill=CREAM, width=lw)         # côté droit
pd.line([(pcx-pw//3, pcy+ph//2), (pcx+pw//3, pcy+ph//2)], fill=CREAM, width=lw)         # fond
pd.line([(pcx-pw//2, pcy-ph//3), (pcx-pw//2-int(pw*0.18), pcy-ph)], fill=CREAM, width=lw)  # poignée
pd.ellipse([pcx-pw//4-6*S, pcy+ph//2+4*S-lw, pcx-pw//4+6*S, pcy+ph//2+4*S+lw], fill=GOLD)
pd.ellipse([pcx+pw//4-6*S, pcy+ph//2+4*S-lw, pcx+pw//4+6*S, pcy+ph//2+4*S+lw], fill=GOLD)

# --- Lettre TS au centre dans un médaillon ---
med_r = int(R*0.34)
d.ellipse([cx-med_r, cy-med_r, cx+med_r, cy+med_r], fill=CREAM, outline=GOLD, width=int(med_r*0.14))
try:
    from PIL import ImageFont
    font = None
    for fp in ("arialbd.ttf", "segoeuib.ttf", "DejaVuSans-Bold.ttf"):
        try:
            font = ImageFont.truetype(fp, int(med_r*0.95)); break
        except Exception: pass
    if font is None:
        font = ImageFont.load_default()
    txt = "TS"
    bbox = d.textbbox((0,0), txt, font=font)
    tw_, th_ = bbox[2]-bbox[0], bbox[3]-bbox[1]
    d.text((cx - tw_//2 - bbox[0], cy - th_//2 - bbox[1]), txt, font=font, fill=BROWN)
except Exception as e:
    print("font fail:", e)

final = img.resize((N, N), Image.LANCZOS)
final.save("/workspace/icon.png")
final.save("/workspace/icon.ico", sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
print("icon generated OK")

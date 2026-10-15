"""
Pictures for PowerPoint decks.

deck_images(title, theme, mode) -> {'cover': path, 'closing': path, 'section': path}
  mode 'ai'  : AI-generated pictures from the free Pollinations service (no key needed);
               falls back to built-in artwork when the service cannot be reached.
  mode 'art' : built-in artwork drawn here (telecom tower, radio waves, network), offline.
Only a generic prompt (topic and colours) is sent to the image service, never any data.
"""
import hashlib, math, os, random, tempfile, urllib.parse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

CACHE = Path(tempfile.gettempdir()) / 'rfai_ppt_images'
SIZES = {'cover': (1210, 1350), 'closing': (707, 1350), 'section': (810, 1350)}   # match the slide areas (~96 dpi x1.4)
PALETTE = {   # (dark, mid, accent, glow)
    'airtel': ('5A0008', 'B0000F', 'FF2A2A', 'FFD6D9'),
    'midnight': ('0B0C11', '1C1F2A', 'E40000', 'FF8A8A'),
    'clean': ('F2F3F6', 'E6E8EE', 'E40000', 'FFFFFF'),
    'ocean': ('061F3D', '0B4C8F', '3AA0FF', 'CFE8FF'),
}
PROMPTS = {
    'cover': 'cinematic photo of a modern 5G telecom tower at dusk, {tone} light, city skyline, clean, high detail, no text, no logo',
    'closing': 'abstract glowing network of connected lights over a city at night, {tone} tones, minimal, no text, no logo',
    'section': 'telecom engineers installing antennas on a rooftop tower, {tone} tones, professional photo, no text, no logo',
}
TONE = {'airtel': 'warm red and white', 'midnight': 'dark charcoal with red accents', 'clean': 'bright white and soft red',
        'ocean': 'deep blue and cyan'}


def deck_images(title, theme, mode):
    CACHE.mkdir(parents=True, exist_ok=True)
    out = {}
    for slot, size in SIZES.items():
        p = None
        if mode == 'ai':
            p = _ai(slot, theme, size)
        if p is None:
            p = _art(slot, theme, size)
        out[slot] = p
    return out


def _ai(slot, theme, size):
    prompt = PROMPTS[slot].format(tone=TONE.get(theme, TONE['airtel']))
    key = hashlib.md5(f'{prompt}{size}'.encode()).hexdigest()
    path = CACHE / f'ai_{key}.jpg'
    if path.exists(): return str(path)
    try:
        import requests
        url = ('https://image.pollinations.ai/prompt/' + urllib.parse.quote(prompt) +
               f'?width={size[0]}&height={size[1]}&nologo=true&seed={int(key[:6], 16) % 100000}')
        r = requests.get(url, timeout=float(os.environ.get('AI_IMAGE_TIMEOUT', '45')))
        if r.status_code != 200 or not r.headers.get('content-type', '').startswith('image'): return None
        raw = CACHE / f'ai_{key}.raw'; raw.write_bytes(r.content)
        im = Image.open(raw).convert('RGB'); raw.unlink()
        _fit(im, size).save(path, quality=88)
        return str(path)
    except Exception:
        return None


def _fit(im, size):
    """Crop to the slot's shape, then resize."""
    tw, th = size; w, h = im.size
    if w / h > tw / th:
        nw = int(h * tw / th); im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else:
        nh = int(w * th / tw); im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    return im.resize(size, Image.LANCZOS)


def _hex(h, a=255): return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


def _art(slot, theme, size):
    """Built-in artwork: gradient sky, radio waves from a lattice tower, a network of glowing nodes."""
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / f'art_{theme}_{slot}_{size[0]}x{size[1]}.png'
    if path.exists(): return str(path)
    dark, mid, acc, glow = PALETTE.get(theme, PALETTE['airtel'])
    light = theme == 'clean'
    S = 2; W, H = size[0] * S, size[1] * S
    rnd = random.Random(f'{theme}{slot}')
    im = Image.new('RGBA', (W, H))
    d = ImageDraw.Draw(im)
    a, b = _hex(dark), _hex(mid)
    for y in range(H):                      # vertical gradient
        t = y / H
        d.line([(0, y), (W, y)], fill=tuple(int(a[i] + (b[i] - a[i]) * (1 - t) ** 1.4) for i in range(3)) + (255,))
    glow_l = Image.new('RGBA', (W, H)); g = ImageDraw.Draw(glow_l)
    tx, base = int(W * (0.55 if slot == 'cover' else 0.5)), int(H * 0.92)
    top = int(H * (0.30 if slot == 'cover' else 0.36))
    cx, cy = tx, top
    # radio waves
    for i in range(1, 9):
        r = int(W * 0.09 * i)
        al = max(30, 170 - i * 17)
        for start, end in ((150, 210), (-30, 30)):        # left and right fans
            g.arc([cx - r, cy - r, cx + r, cy + r], start, end, fill=_hex(glow if not light else acc, al), width=int(5 * S))
    # network nodes
    pts = [(rnd.randint(0, W), rnd.randint(int(H * 0.05), int(H * 0.85))) for _ in range(26)]
    for i, p in enumerate(pts):
        for q in pts[i + 1:i + 3]:
            if math.dist(p, q) < W * 0.45:
                g.line([p, q], fill=_hex(glow if not light else acc, 55), width=2 * S)
    for p in pts:
        rr = rnd.randint(3, 7) * S
        g.ellipse([p[0] - rr, p[1] - rr, p[0] + rr, p[1] + rr], fill=_hex(acc if rnd.random() < .4 else glow, 200))
    blur = glow_l.filter(ImageFilter.GaussianBlur(10 * S))
    im = Image.alpha_composite(im, blur); im = Image.alpha_composite(im, glow_l)
    d = ImageDraw.Draw(im)
    # lattice tower
    col = _hex('FFFFFF' if not light else '2B2F3A', 235)
    half_b, half_t = int(W * 0.13), int(W * 0.018)
    lx = lambda y, s: cx + s * (half_t + (half_b - half_t) * (y - top) / (base - top))
    d.line([(lx(top, -1), top), (lx(base, -1), base)], fill=col, width=6 * S)
    d.line([(lx(top, 1), top), (lx(base, 1), base)], fill=col, width=6 * S)
    ys = [top + (base - top) * (k / 9) ** 1.15 for k in range(10)]
    for y0, y1 in zip(ys, ys[1:]):
        d.line([(lx(y0, -1), y0), (lx(y0, 1), y0)], fill=col, width=3 * S)
        d.line([(lx(y0, -1), y0), (lx(y1, 1), y1)], fill=col, width=2 * S)
        d.line([(lx(y0, 1), y0), (lx(y1, -1), y1)], fill=col, width=2 * S)
    # antenna panels and beacon
    for s_ in (-1, 1):
        d.rounded_rectangle([cx + s_ * int(W * 0.045) - 5 * S, top + 10 * S, cx + s_ * int(W * 0.045) + 5 * S, top + 70 * S], 3 * S, fill=col)
    d.line([(cx, top), (cx, top - 60 * S)], fill=col, width=4 * S)
    bl = Image.new('RGBA', (W, H)); ImageDraw.Draw(bl).ellipse([cx - 26 * S, top - 86 * S, cx + 26 * S, top - 34 * S], fill=_hex(acc, 255))
    im = Image.alpha_composite(im, bl.filter(ImageFilter.GaussianBlur(12 * S)))
    ImageDraw.Draw(im).ellipse([cx - 9 * S, top - 69 * S, cx + 9 * S, top - 51 * S], fill=_hex('FFFFFF', 255))
    # ground glow
    gl = Image.new('RGBA', (W, H)); ImageDraw.Draw(gl).ellipse([cx - W, base - 40 * S, cx + W, base + H * 0.4], fill=_hex(acc, 120))
    im = Image.alpha_composite(im, gl.filter(ImageFilter.GaussianBlur(40 * S)))
    im.convert('RGB').resize(size, Image.LANCZOS).save(path)
    return str(path)

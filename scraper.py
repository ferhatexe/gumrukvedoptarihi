"""
HTTP-based Customs Declaration Scraper.
No Selenium / Chrome required. Uses pure urllib HTTP requests + pixel-based CAPTCHA solving.
"""
import re
import base64
import time
import urllib.request
import urllib.parse
import http.cookiejar
import ssl
import threading
import io
import os
import json
from PIL import Image

# ─────────────────────────────────────────────────────────────
# CAPTCHA Solver — position-based parsing + normalized bitmap matching
# ─────────────────────────────────────────────────────────────
# The CAPTCHA format is ALWAYS: DD + D = ?
#   - Two digits (first number, 10-99)
#   - A '+' operator at a fixed position range (x ≈ 33-47)
#   - One digit (second number, 0-9)
#   - '=' and '?' on the right side (ignored)
# Strategy:
#   1. Binarize, filter colored noise (keep only dark achromatic pixels)
#   2. Remove small connected components (circle-edge noise)
#   3. Column-projection segmentation on left 55% of image
#   4. Split segments into: before operator zone | operator zone | after operator zone
#   5. Classify only the digit segments (never try to classify '+')
#   6. Compute answer = first_number + second_number

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _is_text_pixel(r, g, b):
    """Dark and achromatic = text. Colored = noise circle."""
    brightness = (r + g + b) / 3
    spread = max(r, g, b) - min(r, g, b)
    return brightness < 145 and spread < 25

def _binarize(img_data):
    img = Image.open(io.BytesIO(img_data)).convert("RGB")
    w, h = img.size
    pixels = img.load()
    binary = [[_is_text_pixel(*pixels[x, y]) for x in range(w)] for y in range(h)]
    return binary, w, h

def _remove_small_components(binary, w, h, min_size=10):
    """Remove connected components with fewer than min_size pixels (noise)."""
    visited = [[False]*w for _ in range(h)]
    cleaned = [[False]*w for _ in range(h)]
    for y in range(h):
        for x in range(w):
            if binary[y][x] and not visited[y][x]:
                comp = []
                queue = [(y, x)]
                visited[y][x] = True
                while queue:
                    cy, cx = queue.pop(0)
                    comp.append((cy, cx))
                    for dy in (-1, 0, 1):
                        for dx in (-1, 0, 1):
                            ny, nx = cy + dy, cx + dx
                            if 0 <= ny < h and 0 <= nx < w and binary[ny][nx] and not visited[ny][nx]:
                                visited[ny][nx] = True
                                queue.append((ny, nx))
                if len(comp) >= min_size:
                    for cy, cx in comp:
                        cleaned[cy][cx] = True
    return cleaned

def _col_counts(binary, w, h):
    return [sum(binary[y][x] for y in range(h)) for x in range(w)]

def _find_segments(cc, min_gap=2):
    segs, in_seg, s = [], False, 0
    for i, c in enumerate(cc):
        if c >= 1:
            if not in_seg:
                s = i
                in_seg = True
        else:
            if in_seg:
                segs.append((s, i - 1))
                in_seg = False
    if in_seg:
        segs.append((s, len(cc) - 1))
    merged = []
    for seg in segs:
        if merged and seg[0] - merged[-1][1] <= min_gap:
            merged[-1] = (merged[-1][0], seg[1])
        else:
            merged.append(seg)
    return merged

def _extract_bitmap(binary, h, xs, xe):
    ymin, ymax = h, 0
    for y in range(h):
        for x in range(xs, xe + 1):
            if binary[y][x]:
                ymin = min(ymin, y)
                ymax = max(ymax, y)
                break
    if ymin > ymax:
        return None, 0, 0
    bmp = [[binary[y][x] for x in range(xs, xe + 1)] for y in range(ymin, ymax + 1)]
    return bmp, xe - xs + 1, ymax - ymin + 1

def _normalize(bmp, sw, sh, tw=12, th=20):
    if sh == 0 or sw == 0:
        return [[0] * tw for _ in range(th)]
    result = []
    for ty in range(th):
        sy = min(int(ty * sh / th), sh - 1)
        row = []
        for tx in range(tw):
            sx = min(int(tx * sw / tw), sw - 1)
            row.append(1 if bmp[sy][sx] else 0)
        result.append(row)
    return result

def _bitmap_distance(b1, b2):
    diff = 0
    total = 0
    for r1, r2 in zip(b1, b2):
        for v1, v2 in zip(r1, r2):
            if v1 != v2:
                diff += 1
            total += 1
    return diff / total if total > 0 else 1.0

def _split_wide_segment(binary, h, xs, xe, full_cc):
    """Split a wide segment at column-count valley."""
    sw = xe - xs + 1
    seg_cols = full_cc[xs:xe + 1]
    quarter = sw // 4
    search_s = max(3, quarter)
    search_e = min(sw - 3, sw - quarter)
    min_val, min_pos = float('inf'), sw // 2
    for i in range(search_s, search_e):
        if seg_cols[i] < min_val:
            min_val = seg_cols[i]
            min_pos = i
    split = xs + min_pos
    return [(xs, split - 1), (split, xe)]

# ─── Digit reference database ───
_DIGIT_REFS = None
_REFS_LOCK = threading.Lock()

def _load_digit_refs():
    global _DIGIT_REFS
    with _REFS_LOCK:
        if _DIGIT_REFS is not None:
            return
        ref_path = os.path.join(_SCRIPT_DIR, "digit_references.json")
        if os.path.exists(ref_path):
            with open(ref_path, "r") as f:
                raw = json.load(f)
            # Only load digit references (0-9), NOT '+'
            _DIGIT_REFS = {}
            for char, samples in raw.items():
                if char.isdigit():
                    _DIGIT_REFS[char] = [[[bool(v) for v in row] for row in s] for s in samples]
        else:
            _DIGIT_REFS = {}

def _classify_digit(bmp, sw, sh):
    """Classify a bitmap as a digit 0-9. Returns (char, distance)."""
    _load_digit_refs()
    if not _DIGIT_REFS:
        return '?', 1.0
    norm = _normalize(bmp, sw, sh)
    best_char, best_dist = '?', float('inf')
    for char, samples in _DIGIT_REFS.items():
        for sample in samples:
            dist = _bitmap_distance(norm, sample)
            if dist < best_dist:
                best_dist = dist
                best_char = char
    return best_char, best_dist

def _is_cross_shaped(bmp, sw, sh):
    """Check if a bitmap has a cross/plus shape (horizontal bar in middle, vertical bar)."""
    if sh < 4 or sw < 4:
        return False
    mid_y = sh // 2
    # Middle row should be mostly filled
    mid_count = sum(1 for x in range(sw) if bmp[mid_y][x])
    # Top and bottom rows should be narrow (just the vertical bar)
    top_count = sum(1 for x in range(sw) if bmp[min(1, sh - 1)][x])
    bot_count = sum(1 for x in range(sw) if bmp[max(0, sh - 2)][x])
    avg_edge = max(top_count, bot_count)
    return mid_count > avg_edge * 1.5 and mid_count > sw * 0.4

def solve_captcha(img_data):
    """
    Solve CAPTCHA. Returns (answer_str, expression_str) or (None, error_str).
    Format: DD + D = ? → answer = DD + D (always addition, result 0-200).
    """
    try:
        binary, w, h = _binarize(img_data)
        cleaned = _remove_small_components(binary, w, h, min_size=10)

        # Only look at left 55% (before '=' and '?')
        cutoff = int(w * 0.55)
        cc = _col_counts(cleaned, cutoff, h)
        full_cc = _col_counts(cleaned, w, h)
        raw_segs = _find_segments(cc, min_gap=2)

        if not raw_segs:
            return None, "no segments"

        # Split wide segments (likely two merged digits)
        all_segs = []
        for s, e in raw_segs:
            sw = e - s + 1
            if sw >= 18:
                parts = _split_wide_segment(cleaned, h, s, e, full_cc)
                all_segs.extend(parts)
            else:
                all_segs.append((s, e))

        # ── Find the '+' operator ──
        # Strategy: look for the segment that (a) is closest to x ≈ 30% of width
        # and (b) has a cross shape. If no cross found, use position only.
        expected_op_x = w * 0.30  # '+' is typically around 30% of image width
        op_zone_min = w * 0.20
        op_zone_max = w * 0.45

        best_op_idx = -1
        best_op_score = float('inf')

        for i, (s, e) in enumerate(all_segs):
            mid = (s + e) / 2
            if mid < op_zone_min or mid > op_zone_max:
                continue
            bmp, sw, sh = _extract_bitmap(cleaned, h, s, e)
            if bmp is None:
                continue
            
            # Score: distance from expected position + bonus for cross shape
            pos_score = abs(mid - expected_op_x)
            if _is_cross_shaped(bmp, sw, sh):
                pos_score -= 20  # strong bonus for cross shape
            
            if pos_score < best_op_score:
                best_op_score = pos_score
                best_op_idx = i

        if best_op_idx < 0:
            # Fallback: use pure position (segment closest to 30%)
            for i, (s, e) in enumerate(all_segs):
                mid = (s + e) / 2
                score = abs(mid - expected_op_x)
                if score < best_op_score:
                    best_op_score = score
                    best_op_idx = i

        # Partition: everything before operator = first number, after = second number
        left_segs = all_segs[:best_op_idx]
        right_segs = all_segs[best_op_idx + 1:]

        # Extract and classify the operator
        op_char = '+'
        if 0 <= best_op_idx < len(all_segs):
            op_s, op_e = all_segs[best_op_idx]
            op_bmp, op_sw, op_sh = _extract_bitmap(cleaned, h, op_s, op_e)
            if op_bmp is not None:
                if op_sh <= 3 or op_sh < op_sw * 0.5:
                    op_char = '-'
                else:
                    op_char = '+'

        # Classify left segments as digits
        left_digits = []
        for s, e in left_segs:
            bmp, sw, sh = _extract_bitmap(cleaned, h, s, e)
            if bmp is None or sh < 3:
                continue
            char, dist = _classify_digit(bmp, sw, sh)
            if dist < 0.45:
                left_digits.append(char)

        # Classify right segments as digits
        right_digits = []
        for s, e in right_segs:
            bmp, sw, sh = _extract_bitmap(cleaned, h, s, e)
            if bmp is None or sh < 3:
                continue
            char, dist = _classify_digit(bmp, sw, sh)
            if dist < 0.45:
                right_digits.append(char)

        if not left_digits or not right_digits:
            return None, f"L={''.join(left_digits)} R={''.join(right_digits)}"

        left_str = "".join(left_digits)
        right_str = "".join(right_digits)
        expr = f"{left_str}{op_char}{right_str}"

        try:
            n1 = int(left_str)
            n2 = int(right_str)
        except ValueError:
            return None, expr

        if op_char == '-':
            result = n1 - n2
        else:
            result = n1 + n2

        if -100 <= result <= 200:
            return str(result), expr
        return None, f"{expr}={result} out of range"

    except Exception as e:
        return None, f"Error: {e}"


# ─────────────────────────────────────────────────────────────
# HTTP Scraper
# ─────────────────────────────────────────────────────────────

class HttpCustomsScraper:
    """
    Queries the Turkish Customs portal (uygulama.gtb.gov.tr) via HTTP.
    Supports both standard GCB (18-char) and ETGB (16-char) declarations.
    """
    URL = "https://uygulama.gtb.gov.tr/beyannamesorgulama/"

    def __init__(self, log_callback=None, cancel_check=None):
        self.log_callback = log_callback
        self.cancel_check = cancel_check
        self.ssl_ctx = ssl.create_default_context()
        self.ssl_ctx.check_hostname = False
        self.ssl_ctx.verify_mode = ssl.CERT_NONE

    def log(self, message):
        try:
            print(message)
        except UnicodeEncodeError:
            print(str(message).encode('ascii', errors='replace').decode('ascii'))
        if self.log_callback:
            try:
                self.log_callback(message)
            except Exception:
                pass

    def is_cancelled(self):
        return self.cancel_check() if self.cancel_check else False

    def _create_opener(self):
        cj = http.cookiejar.CookieJar()
        return urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(cj),
            urllib.request.HTTPSHandler(context=self.ssl_ctx)
        )

    def _make_request(self, opener, url, data=None):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        }
        if data is not None:
            headers['Content-Type'] = 'application/x-www-form-urlencoded'
            headers['Referer'] = url
            encoded = urllib.parse.urlencode(data).encode('utf-8')
            req = urllib.request.Request(url, data=encoded, headers=headers)
        else:
            req = urllib.request.Request(url, headers=headers)
        resp = opener.open(req, timeout=15)
        return resp.read().decode('utf-8')

    def _parse_form_fields(self, html):
        vs = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', html)
        ev = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', html)
        vsg = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', html)
        return {
            'viewstate': vs.group(1) if vs else '',
            'eventvalidation': ev.group(1) if ev else '',
            'vsg': vsg.group(1) if vsg else ''
        }

    def _extract_captcha_image(self, html):
        img_match = re.search(r'id="imgGuvenlik"\s+src="data:image/png;base64,([^"]+)"', html)
        if not img_match:
            return None
        return base64.b64decode(img_match.group(1))

    def _do_etgb_postback(self, opener, fields):
        post_data = {
            '__EVENTTARGET': 'rdETGB', '__EVENTARGUMENT': '', '__LASTFOCUS': '',
            '__VIEWSTATE': fields['viewstate'],
            '__VIEWSTATEGENERATOR': fields['vsg'],
            '__EVENTVALIDATION': fields['eventvalidation'],
            'Sorgu': 'rdETGB', 'txtBeyannameNo': '',
            'txtDogrulamaKodu': '', 'txtModalPopUpGoster': '0'
        }
        html = self._make_request(opener, self.URL, data=post_data)
        return html, self._parse_form_fields(html)

    def _parse_result(self, html):
        durum_match = re.search(r'id="lblBeyannameDurum"[^>]*>(.*?)</span>', html, re.DOTALL)
        error_match = re.search(r'id="LabelDurum"[^>]*>(.*?)</span>', html, re.DOTALL)
        durum_val = durum_match.group(1).strip() if durum_match else ''
        error_val = error_match.group(1).strip() if error_match else ''

        if durum_val:
            durum_upper = durum_val.upper()
            if 'KAPANMAMIŞ' in durum_upper or 'KAPANMAMI' in durum_upper:
                return {"success": True, "status": "Kapanmamış", "message": "Beyanname kapanmamış.", "date": None}
            date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', durum_val)
            if date_match:
                d, m, y = date_match.group(1), date_match.group(2), date_match.group(3)
                return {"success": True, "status": "İntaç Tarihi Var", "message": "İntaç tarihi bulundu.", "date": f"{y}-{m}-{d}"}
            if 'DOLMAMI' in durum_upper or 'SÜRES' in durum_upper or 'SURES' in durum_upper:
                return {"success": False, "status": "RateLimit", "message": durum_val, "date": None}
            return {"success": True, "status": "Tarih Okunamadı", "message": durum_val, "date": None}

        if error_val:
            error_upper = error_val.upper()
            if any(kw in error_upper for kw in ['GÜVENL', 'KODU', 'DOĞRULAMA', 'YANLI']):
                return {"success": False, "status": "CaptchaWrong", "message": error_val, "date": None}
            return {"success": False, "status": "Sistem Uyarısı", "message": error_val, "date": None}

        return {"success": False, "status": "Bilinmeyen", "message": "Sonuç ayrıştırılamadı.", "date": None}

    def _interruptible_sleep(self, seconds, log_msg_prefix=None):
        """Sleep for `seconds` but check cancel every 1s. Returns True if cancelled."""
        steps = int(seconds)
        for i in range(steps):
            if self.is_cancelled():
                return True
            if log_msg_prefix and i > 0 and i % 10 == 0:
                remaining = seconds - i
                rem_min = int(remaining // 60)
                rem_sec = int(remaining % 60)
                self.log(f"{log_msg_prefix} ({rem_min}dk {rem_sec}sn kaldı...)")
            time.sleep(1)
        # Fractional remainder
        remainder = seconds - steps
        if remainder > 0:
            time.sleep(remainder)
        return self.is_cancelled()

    def _parse_wait_seconds(self, message):
        """Parse wait time from rate-limit message. Returns seconds to wait."""
        # Try to find "X dakika Y saniye" or similar patterns
        msg_upper = message.upper()
        minutes = 0
        seconds = 0
        m = re.search(r'(\d+)\s*DAK', msg_upper)
        if m:
            minutes = int(m.group(1))
        s = re.search(r'(\d+)\s*SAN', msg_upper)
        if s:
            seconds = int(s.group(1))
        total = minutes * 60 + seconds
        if total > 0:
            return min(total + 5, 600)  # Add 5s buffer, cap at 10 min
        return 310  # Default: 5 min 10 sec if can't parse

    def query_declaration(self, gcb_no, max_attempts=None):
        """
        Query a single declaration by GCB number.
        Retries INDEFINITELY until a finalized result is obtained:
          - "İntaç Tarihi Var" (with a date)
          - "Kapanmamış" (declaration not closed)
        Only stops on cancel or finalized result.
        Rate-limit cooldowns (up to 5+ minutes) are waited out automatically.
        """
        gcb_no = gcb_no.strip().upper()
        is_etgb = len(gcb_no) == 16
        attempt = 0

        while True:
            attempt += 1
            if self.is_cancelled():
                return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}
            try:
                opener = self._create_opener()
                html = self._make_request(opener, self.URL)
                if self.is_cancelled():
                    return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}
                fields = self._parse_form_fields(html)

                if is_etgb:
                    html, fields = self._do_etgb_postback(opener, fields)

                img_data = self._extract_captcha_image(html)
                if not img_data:
                    continue

                solution, ocr_text = solve_captcha(img_data)
                if not solution:
                    self.log(f"[{gcb_no}] Deneme {attempt}: OCR çözülemedi ('{ocr_text}')")
                    continue

                if self.is_cancelled():
                    return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}

                post_data = {
                    '__EVENTTARGET': '', '__EVENTARGUMENT': '', '__LASTFOCUS': '',
                    '__VIEWSTATE': fields['viewstate'],
                    '__VIEWSTATEGENERATOR': fields['vsg'],
                    '__EVENTVALIDATION': fields['eventvalidation'],
                    'txtBeyannameNo': gcb_no, 'txtDogrulamaKodu': solution,
                    'txtModalPopUpGoster': '0', 'btnSorgula': 'Sorgula'
                }
                if is_etgb:
                    post_data['Sorgu'] = 'rdETGB'
                    post_data['CheckBoxETGB'] = 'on'
                else:
                    post_data['Sorgu'] = 'rdBeyanname'

                resp_html = self._make_request(opener, self.URL, data=post_data)
                result = self._parse_result(resp_html)

                # ── Finalized results: return immediately ──
                if result['status'] == 'İntaç Tarihi Var':
                    return result
                if result['status'] == 'Kapanmamış':
                    return result

                # ── Retryable results: keep going ──
                if result['status'] == 'CaptchaWrong':
                    self.log(f"[{gcb_no}] Deneme {attempt}: Güvenlik kodu yanlış ('{ocr_text}' -> {solution})")
                    continue

                if result['status'] == 'RateLimit':
                    # Return immediately without sleeping, let the client handle the cooldown
                    return result

                # Non-finalized results (Tarih Okunamadı, Sistem Uyarısı, Bilinmeyen etc.)
                # Log and retry
                self.log(f"[{gcb_no}] Deneme {attempt}: {result.get('status', '?')}: {result.get('message', '?')} — tekrar deneniyor...")
                if self._interruptible_sleep(2):
                    return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}
                continue

            except Exception as e:
                if self.is_cancelled():
                    return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}
                err_msg = str(e).encode('ascii', errors='replace').decode('ascii') if str(e) else 'Bilinmeyen hata'
                self.log(f"[{gcb_no}] Deneme {attempt}: Hata: {err_msg}")
                if self._interruptible_sleep(1):
                    return {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}
                continue

    def close(self):
        pass


if __name__ == "__main__":
    import sys
    gcb = "26341200EX00137190"
    if len(sys.argv) > 1:
        gcb = sys.argv[1]
    scraper = HttpCustomsScraper()
    result = scraper.query_declaration(gcb)
    print(f"\nResult: {result}")

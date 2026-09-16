# -*- coding: utf-8 -*-
"""2026년 4회 기출복원 ↔ 교재 2권 문항 크롭 갤러리 빌더."""
import sys, io, os, re, ast, json, difflib, html as H
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import fitz, openpyxl
from PIL import Image

SRC  = r'D:\electrician-cbt\_build_wongo_snapshot.py'
PDF  = r'D:\CBT 모의고사\기출\[2027 전기기능사 필기(2권)]-최종_20260728.pdf'
MST  = r'D:\electrician-2027-0721-review\마스터_1권2권_전체구조_인수인계_2026-07-22_v3.xlsx'
CBT  = r'D:\electrician-cbt\전기기능사_CBT_31회차_통합.xlsx'
OUTD = r'D:\electrician-cbt\_crop_2026-4회'
OUTH = r'D:\electrician-cbt\_2026-4회_교재매칭_크롭.html'
RELD = '_crop_2026-4회'

DPI  = 180
ZOOM = DPI / 72.0
COLW = 191.1          # 단 폭
PITCH = 216.8         # 단 피치
PADX = 9.0
PADY = 7.0
COL_TOP = 48.0        # 단 상단(연속 세그먼트용)
COL_BOT = 688.0       # 단 하단(푸터 위)
GAP_BREAK = 58.0      # 문항 내부 최대 세로 공백

SUBJ = {'이론': '전기이론', '기기': '전기기기', '설비': '전기설비'}
TIERNOTE = {
 'O': '교재 2권 미수록 — 2002~2009 구기출',
 'R': '교재 2권 미수록 — 시험장 신규 복원',
 'T': '교재 2권(기출) 미수록 — 1권 이론에 수록',
 'N': '교재 2권 미수록 — 전문 복원 필요',
 'X': '교재 매칭 제외 — 대표 출처 철회(등급 X)',
}
TIERLABEL = {'A': 'A · DB 원문', 'W': 'C · 재작성', 'O': 'A · 구기출',
             'R': '신규 복원', 'T': '1권 이론', 'N': '신규'}

# ── ITEMS 추출 ─────────────────────────────────────────
src = open(SRC, encoding='utf-8').read()
tree = ast.parse(src)
ITEMS = None
for node in tree.body:
    if isinstance(node, ast.Assign) and getattr(node.targets[0], 'id', '') == 'ITEMS':
        ITEMS = ast.literal_eval(node.value)
assert ITEMS, 'ITEMS 추출 실패'
print('ITEMS 총', len(ITEMS))

# ── 마스터: (연도,회차,번호) -> (인쇄쪽, 과목, 소단원, 별표) ────
page_of, star_of, unit_of = {}, {}, {}
mw = openpyxl.load_workbook(MST, read_only=True, data_only=True)
for r in mw['2권 문항마스터'].iter_rows(min_row=2, values_only=True):
    try:
        y, h = str(r[0]).split('-')
        k = (int(y), int(h), int(r[1]))
    except Exception:
        continue
    page_of[k] = r[2]
    unit_of[k] = r[4] or ''
    star_of[k] = r[6] or ''
print('마스터 문항', len(page_of))

# ── CBT DB 발문(검증용) ────────────────────────────────
REWRITE = {}
for node in tree.body:
    if isinstance(node, ast.Assign) and getattr(node.targets[0], 'id', '') == 'REWRITE':
        REWRITE = ast.literal_eval(node.value)
print('REWRITE', len(REWRITE))

bank_stem, bank_row = {}, {}
wb = openpyxl.load_workbook(CBT, read_only=True, data_only=True)
ws = wb[wb.sheetnames[0]]
rows = list(ws.iter_rows(min_row=2, values_only=True))
hdr = [str(x) if x else '' for x in rows[0]]
idx_c = {h: i for i, h in enumerate(hdr)}
for r in rows[1:]:
    try:
        k = (int(float(str(r[idx_c['연도']]))), int(float(str(r[idx_c['회차']]))),
             int(float(str(r[idx_c['번호']]))))
    except Exception:
        continue
    bank_stem[k] = str(r[idx_c['발문']] or '')
    bank_row[k] = r

RESTORED, OLD = {}, {}
for node in tree.body:
    if isinstance(node, ast.Assign):
        nm = getattr(node.targets[0], 'id', '')
        if nm == 'RESTORED':
            RESTORED = ast.literal_eval(node.value)
        elif nm == 'OLD':
            OLD = ast.literal_eval(node.value)
print('RESTORED', len(RESTORED), '/ OLD', len(OLD))


def fixsrc(s, y, h):
    """DB 발문의 <img src="..."> 를 회차 폴더 상대경로로 되돌린다(원본 build_wongo 규칙)."""
    if 'src="' not in s:
        return s
    pre = 'CBT_%d_%d회/img/' % (y, h)
    return s.replace('src="', 'src="' + pre).replace(pre + 'CBT_', 'CBT_')


def exam_item(tier, srcv):
    """복원 문항 전문 → (발문HTML, 보기4, 정답번호). 없으면 None."""
    try:
        if tier in ('W', 'X') and srcv in REWRITE:
            return REWRITE[srcv]
        if tier in ('R', 'T') and srcv in RESTORED:
            return RESTORED[srcv]
        if tier == 'O' and srcv in OLD:
            return OLD[srcv]
        if isinstance(srcv, tuple) and len(srcv) == 3 and srcv in bank_row:
            r = bank_row[srcv]
            y, h, _ = srcv
            g = lambda k: str(r[idx_c[k]] or '')
            a = g('정답(1~4)').strip()
            return (fixsrc(g('발문'), y, h),
                    [fixsrc(g('보기%d' % k), y, h) for k in range(1, 5)],
                    int(a) if a.isdigit() else 0)
    except Exception:
        pass
    return None


HAN = re.compile(r'[가-힣]+')
def hangul(s):
    s = re.sub(r'<[^>]+>', ' ', s or '')
    return ''.join(HAN.findall(s))

# ── PDF ────────────────────────────────────────────────
doc = fitz.open(PDF)
NPAGE = doc.page_count

_bcache = {}
def page_bands(pi):
    """단을 가로지르는 페이지 장식(회차 표지 띠·빈출도별 구성 박스 등)의 세로 범위.
    2단 조판에서 문항 콘텐츠는 한 단(191pt)을 넘지 않으므로,
    폭 200pt를 넘는 색면·이미지는 문항이 아니라 페이지 장식이다."""
    if pi in _bcache:
        return _bcache[pi]
    p = doc[pi]
    bs = []
    for dr in p.get_drawings():
        r = dr['rect']
        if r.width > 200 and r.height > 6 and r.y0 < COL_BOT:
            bs.append([r.y0, r.y1])
    for im in p.get_image_info():
        x0, y0, x1, y1 = im['bbox']
        if x1 - x0 > 200 and y0 < COL_BOT:
            bs.append([y0, y1])
    bs.sort()
    merged = []
    for b in bs:
        if merged and b[0] <= merged[-1][1] + 6:
            merged[-1][1] = max(merged[-1][1], b[1])
        else:
            merged.append(list(b))
    # 띠 위에 얹힌 글자(제목·「빈출도별 구성」 수치)까지 띠에 흡수
    if merged:
        for b in p.get_text("dict")['blocks']:
            if b['type'] != 0:
                continue
            for ln in b['lines']:
                for s in ln['spans']:
                    if not s['text'].strip():
                        continue
                    y0, y1 = s['bbox'][1], s['bbox'][3]
                    for m in merged:
                        if y0 < m[1] + 8 and y1 > m[0] - 8:
                            m[0], m[1] = min(m[0], y0), max(m[1], y1)
        merged = [[m[0] - 3, m[1] + 3] for m in merged]
    _bcache[pi] = merged
    return merged


def in_band(pi, y0, y1):
    h = max(y1 - y0, .1)
    for b0, b1 in page_bands(pi):
        if min(y1, b1) - max(y0, b0) > h * 0.5:
            return True
    return False


def band_gap(pi, a, b):
    """a~b 구간에 낀 띠 높이 합(공백 판정에서 빼 준다)."""
    t = 0.0
    for b0, b1 in page_bands(pi):
        t += max(0.0, min(b, b1) - max(a, b0))
    return t


_cache = {}
def page_info(pi):
    """페이지의 문항번호 마커 목록과 콘텐츠 아이템 목록."""
    if pi in _cache:
        return _cache[pi]
    p = doc[pi]
    d = p.get_text("dict")
    col0 = 85.0 if (pi + 1) % 2 == 1 else 70.9   # 인쇄쪽 = pi+1
    cols = [col0, col0 + PITCH]

    def colof(x):
        for ci, cx in enumerate(cols):
            if cx - 9 <= x <= cx + COLW + 9:
                return ci
        return -1          # 본문 단 밖(단 구분선·가장자리 색인 탭 등)

    markers, items = [], []
    for b in d['blocks']:
        if b['type'] != 0:
            continue
        for ln in b['lines']:
            big = [s for s in ln['spans']
                   if 'DIN2014-ExtraBold' in s['font'] and s['size'] > 13
                   and s['text'].strip().isdigit()]
            if big:
                big.sort(key=lambda s: s['bbox'][0])
                txt = ''.join(s['text'].strip() for s in big)
                bb = (min(s['bbox'][0] for s in big), min(s['bbox'][1] for s in big),
                      max(s['bbox'][2] for s in big), max(s['bbox'][3] for s in big))
                if bb[1] < COL_BOT:
                    markers.append({'n': int(txt), 'txt': txt, 'bbox': bb,
                                    'col': colof(bb[0]), 'page': pi})
            for s in ln['spans']:
                t = s['text'].strip()
                if not t:
                    continue
                x0, y0, x1, y1 = s['bbox']
                if y1 > COL_BOT or y0 < 20:
                    continue
                items.append((colof(x0), y0, y1, t))
    # 이미지·도형도 콘텐츠로 취급
    for im in p.get_image_info():
        x0, y0, x1, y1 = im['bbox']
        if y1 > COL_BOT or (x1 - x0) > 500:
            continue
        items.append((colof(x0), y0, y1, ''))
    def visible(dr):
        f, c = dr.get('fill'), dr.get('color')
        white = lambda v: v is not None and min(v) > 0.93
        if c is None and (f is None or white(f)):
            return False
        if dr.get('fill_opacity', 1) == 0 and dr.get('stroke_opacity', 1) == 0:
            return False
        return True

    for dr in p.get_drawings():
        r = dr['rect']
        if r.y1 > COL_BOT or r.width > 480 or r.width <= 0:
            continue
        if r.height > 420:                      # 단 구분선·페이지 장식
            continue
        if r.width < 1.2 and r.height > 60:     # 세로 괘선
            continue
        if not visible(dr):
            continue
        items.append((colof(r.x0), r.y0, r.y1, ''))
    items = [it for it in items if it[0] >= 0 and not in_band(pi, it[1], it[2])]
    markers = [m for m in markers if m['col'] >= 0
               and not in_band(pi, m['bbox'][1], m['bbox'][3])]
    markers.sort(key=lambda m: (m['col'], m['bbox'][1]))
    res = (markers, items, cols)
    _cache[pi] = res
    return res


def col_extent(pi, col, y_from, anchor_first=False):
    """pi 페이지 col 단에서 y_from 아래로 이어지는 콘텐츠의 끝 y.
    anchor_first=True면 단 맨 위 장식을 건너뛰고 첫 콘텐츠부터 센다."""
    _, items, _ = page_info(pi)
    ys = sorted([(y0, y1) for (c, y0, y1, _t) in items if c == col and y1 > y_from - 2])
    if not ys:
        return min(y_from + PADY, COL_BOT)
    end = ys[0][0] if anchor_first else y_from
    for y0, y1 in ys:
        if y0 - end - band_gap(pi, end, y0) > GAP_BREAK:
            break
        end = max(end, y1)
    return min(end + PADY, COL_BOT)


def next_marker(pi, col, y):
    """읽기 순서상 다음 문항 마커."""
    markers, _, _ = page_info(pi)
    for m in markers:
        if (m['col'], m['bbox'][1]) > (col, y + 1):
            return m
    for q in range(pi + 1, min(pi + 3, NPAGE)):
        mk, _, _ = page_info(q)
        if mk:
            return mk[0]
    return None


def segments(pi, marker):
    """문항 영역 세그먼트: [(page, col, y_top, y_bot), ...]"""
    col, y = marker['col'], marker['bbox'][1]
    nm = next_marker(pi, col, y)
    segs = []
    if nm and nm['page'] == pi and nm['col'] == col:
        bot = min(nm['bbox'][1] - PADY, col_extent(pi, col, y))
        return [(pi, col, y - PADY, bot)]
    # 현재 단 끝까지
    ext = col_extent(pi, col, y)
    segs.append((pi, col, y - PADY, ext))
    # 본문이 단 바닥까지 닿지 않으면 그 단에서 문항이 끝난 것 → 이어붙이지 않는다
    if nm is None or ext < COL_BOT - 40:
        return segs
    cur_p, cur_c = pi, col
    guard = 0
    while guard < 4:
        guard += 1
        if cur_c == 0:
            cur_c = 1
        else:
            cur_c = 0
            cur_p += 1
        if cur_p >= NPAGE:
            break
        if (cur_p, cur_c) == (nm['page'], nm['col']):
            bot = min(nm['bbox'][1] - PADY, col_extent(cur_p, cur_c, COL_TOP, True))
            segs.append((cur_p, cur_c, COL_TOP, bot))
            break
        segs.append((cur_p, cur_c, COL_TOP, col_extent(cur_p, cur_c, COL_TOP, True)))
    return [s for s in segs if s[3] - s[2] > 12]


def cut_bands(pi, y0, y1):
    """세로 구간에서 페이지 장식 띠를 잘라내고 남은 조각들."""
    parts = [(y0, y1)]
    for b0, b1 in page_bands(pi):
        nxt = []
        for a, b in parts:
            if b1 <= a or b0 >= b:
                nxt.append((a, b))
                continue
            if b0 - a > 12:
                nxt.append((a, b0))
            if b - b1 > 12:
                nxt.append((b1, b))
        parts = nxt
    return [p for p in parts if p[1] - p[0] > 12]


def render(segs, outpath, clips_out=None):
    imgs, texts = [], []
    for (pi, col, sy0, sy1) in segs:
        _, _, cols = page_info(pi)
        x0 = cols[col] - PADX
        x1 = cols[col] + COLW + PADX
        for (y0, y1) in cut_bands(pi, max(sy0, 24), min(sy1, COL_BOT + 4)):
            clip = fitz.Rect(x0, y0, x1, y1)
            if clips_out is not None:
                clips_out.append((pi, y0, y1))
            pm = doc[pi].get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), clip=clip, alpha=False)
            imgs.append(Image.frombytes('RGB', (pm.width, pm.height), pm.samples))
            texts.append(doc[pi].get_text("text", clip=clip))
    if not imgs:
        return None, '', 0
    W = max(i.width for i in imgs)
    GAPP = 10 if len(imgs) > 1 else 0
    Ht = sum(i.height for i in imgs) + GAPP * (len(imgs) - 1)
    canvas = Image.new('RGB', (W, Ht), (255, 255, 255))
    yy = 0
    for i in imgs:
        canvas.paste(i, (0, yy))
        yy += i.height + GAPP
    canvas.save(outpath, 'PNG', optimize=True)
    return canvas.size, '\n'.join(texts), len(imgs)


def full_page(pi, outpath):
    clip = fitz.Rect(14, 14, 547, 743)
    pm = doc[pi].get_pixmap(matrix=fitz.Matrix(ZOOM * 0.8, ZOOM * 0.8), clip=clip, alpha=False)
    pm.save(outpath)
    return (pm.width, pm.height), doc[pi].get_text()


def find_marker(printed_page, qno):
    """인쇄쪽 기준 ±4쪽에서 문항번호 마커 찾기. 반환 (pi, marker, 이동쪽수)"""
    base = printed_page - 1          # 인쇄쪽 = PDF index + 1
    for off in (0, 1, -1, 2, -2, 3, -3, 4, -4):
        pi = base + off
        if not (0 <= pi < NPAGE):
            continue
        mk, _, _ = page_info(pi)
        hit = [m for m in mk if m['n'] == qno]
        if len(hit) == 1:
            return pi, hit[0], off
        if len(hit) > 1:
            return pi, hit[0], off
    return None, None, None


os.makedirs(OUTD, exist_ok=True)
for f in os.listdir(OUTD):
    if f.lower().endswith('.png'):
        os.remove(os.path.join(OUTD, f))

results = []
seqno = {'이론': 0, '기기': 0, '설비': 0}
for (subj, title, tier, srcv, mine) in ITEMS:
    seqno[subj] += 1
    n = seqno[subj]
    rec = {'subj': subj, 'sn': SUBJ[subj], 'n': n, 'title': title, 'tier': tier,
           'mine': mine, 'status': '', 'img': '', 'src': '', 'star': '', 'unit': '',
           'page': None, 'note': '', 'verify': ''}
    ei = exam_item(tier, srcv)
    if ei:
        rec['q_stem'], rec['q_opts'], rec['q_ans'] = ei[0], list(ei[1]), ei[2]
    else:
        rec['q_stem'], rec['q_opts'], rec['q_ans'] = '', [], 0
        print('  [본문없음]', SUBJ[subj], n, tier, title[:24])
    if tier not in ('A', 'W') or not (isinstance(srcv, tuple) and len(srcv) == 3):
        rec['status'] = 'skip'
        rec['note'] = TIERNOTE.get(tier, '교재 2권 미수록')
        if tier == 'O':
            rec['src'] = ('%s-%s-%s 구기출 %d번'
                          % (srcv[0][:4], srcv[0][4:6], srcv[0][6:8], srcv[1]))
        elif tier == 'T':
            rec['src'] = '1권 이론 수록'
        elif isinstance(srcv, tuple) and len(srcv) == 3:
            rec['src'] = '%d-%d회 %d번' % srcv
            rec['page'] = page_of.get(srcv)
            rec['star'] = star_of.get(srcv, '')
            rec['unit'] = unit_of.get(srcv, '')
        else:
            rec['src'] = '시험장 복원(신규)'
        results.append(rec)
        continue

    y, h, qn = srcv
    pp = page_of.get(srcv)
    rec['src'] = '%d-%d회 %d번' % (y, h, qn)
    rec['star'] = star_of.get(srcv, '')
    rec['unit'] = unit_of.get(srcv, '')
    if not pp:
        rec['status'] = 'fail'
        rec['note'] = '마스터에 쪽수 없음'
        results.append(rec)
        continue
    rec['page'] = pp
    pi, mk, off = find_marker(int(pp), qn)
    fn = '%s_%02d_%d-%d회%02d번.png' % (SUBJ[subj], n, y, h, qn)
    out = os.path.join(OUTD, fn)
    if mk is None:
        pi = int(pp) - 1
        if 0 <= pi < NPAGE:
            full_page(pi, out)
            rec['status'] = 'page'
            rec['img'] = fn
            rec['note'] = '문항번호 미검출 → 쪽 전체'
        else:
            rec['status'] = 'fail'
            rec['note'] = '쪽 범위 밖'
        results.append(rec)
        continue
    segs = segments(pi, mk)
    clips = []
    size, txt, npc = render(segs, out, clips)
    if size is None:
        full_page(pi, out)
        rec['status'] = 'page'
        rec['img'] = fn
        rec['note'] = '영역 계산 실패 → 쪽 전체'
        results.append(rec)
        continue
    rec['status'] = 'ok'
    rec['img'] = fn
    rec['segs'] = npc
    rec['pdfpage'] = pi + 1
    if off:
        rec['note'] = '마스터 %s쪽 → 실제 %d쪽(%+d)' % (pp, pi + 1, off)
    # 검증 0-a: 크롭 영역이 회차 표지 띠(단을 가로지르는 색면)와 겹치는가 — 기하 검사
    geo = []
    for (cp, cy0, cy1) in clips:
        for (b0, b1) in page_bands(cp):
            if min(cy1, b1) - max(cy0, b0) > 2:
                geo.append('%d쪽 y%.0f~%.0f' % (cp + 1, max(cy0, b0), min(cy1, b1)))
    # 검증 0-b: 러닝헤더·회차 마무리표 텍스트 혼입 — 텍스트 검사
    flat = txt.replace(' ', '')
    DIRT = ['빈출도별구성', 'CBT기출문제', '2027대비독끝', 'PARTC최신',
            '회독(20문항)', '회차별정답수']
    dirty = [d for d in DIRT if d in flat] + geo
    if dirty:
        rec['dirty'] = dirty
        rec['note'] = ((rec['note'] + ' / ' if rec['note'] else '')
                       + '장식 혼입: ' + ','.join(dirty))
    # 검증 1: 크롭 첫 줄의 문항번호
    head = txt.strip().split('\n')[0].strip()
    rec['head'] = head
    if head != str(qn) and head != '%02d' % qn:
        rec['note'] = (rec['note'] + ' / ' if rec['note'] else '') + '크롭 첫 줄 번호 불일치(%s)' % head[:8]
    # 검증 2: 지면 [과목] (n) 소단원 (000p) 줄 ↔ 마스터 소단원
    m = re.search(r'\[(전기이론|전기기기|전기설비)\]\s*(\(\d+\).*?)\s*\(\d+p\)', txt)
    pdf_unit = (m.group(2).strip() if m else '')
    rec['pdf_unit'] = pdf_unit
    mu = re.sub(r'\s+', '', rec['unit'])
    pu = re.sub(r'\s+', '', pdf_unit)
    if mu and pu:
        rec['unit_ok'] = (mu == pu)
        if not rec['unit_ok']:
            rec['note'] = ((rec['note'] + ' / ' if rec['note'] else '')
                           + '소단원 불일치(마스터 %s / 지면 %s)' % (rec['unit'], pdf_unit))
    else:
        rec['unit_ok'] = None
    # 검증 3: DB 발문 ↔ 지면 한글 유사도(크롭 정확도 참고치)
    b = hangul(txt)[:420]

    def sim(a):
        a = a[:60]
        if not a or not b:
            return 0.0
        return max(difflib.SequenceMatcher(None, a, b[i:i + len(a) + 20]).ratio()
                   for i in range(0, max(1, len(b) - len(a)), 4))

    rec['verify'] = '%.2f' % sim(hangul(bank_stem.get(srcv, '')))
    # 검증 4: 시험 복원 발문 ↔ 교재 지면 발문 (W는 재작성본이 시험 원문)
    exam = REWRITE[srcv][0] if (tier == 'W' and srcv in REWRITE) else bank_stem.get(srcv, '')
    rec['exam_sim'] = round(sim(hangul(exam)), 2)
    rec['same_q'] = rec['exam_sim'] >= 0.62
    results.append(rec)

json.dump(results, open(os.path.join(OUTD, '_result.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
ok = sum(1 for r in results if r['status'] == 'ok')
pg = sum(1 for r in results if r['status'] == 'page')
fl = sum(1 for r in results if r['status'] == 'fail')
sk = sum(1 for r in results if r['status'] == 'skip')
print('ok %d / 쪽전체 %d / 실패 %d / 제외(교재미수록) %d' % (ok, pg, fl, sk))
for r in results:
    if r['note'] or r['status'] not in ('ok',):
        print('  [%s] %s %02d %s | %s | %s' % (r['status'], r['sn'], r['n'],
              r['title'][:26], r['src'], r['note']))

# -*- coding: utf-8 -*-
"""_result.json → 모바일·다크모드 대응 HTML 갤러리 (완벽적중 / 유형 일치)."""
import sys, io, os, re, json, html as H, datetime
from collections import Counter
import openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

OUTD = r'D:\electrician-cbt\_crop_2026-4회'
OUTH = r'D:\electrician-cbt\_2026-4회_교재매칭_크롭.html'
RELD = '_crop_2026-4회'

R = json.load(open(os.path.join(OUTD, '_result.json'), encoding='utf-8'))
SUBJS = ['전기이론', '전기기기', '전기설비']
SID = {'전기이론': 'th', '전기기기': 'ma', '전기설비': 'fa'}

ok = [r for r in R if r['status'] == 'ok']
pgs = [r for r in R if r['status'] == 'page']
fls = [r for r in R if r['status'] == 'fail']
sks = [r for r in R if r['status'] == 'skip']
shot = ok + pgs
A = [r for r in shot if r['tier'] == 'A']
W = [r for r in shot if r['tier'] == 'W']
MISS = sks + fls
HITN, TOTN = len(shot), len(R)
RATE = HITN / TOTN * 100

# ── 집계: 별점 · 회차 ───────────────────────────────────
STARS = ['★★★', '★★☆', '★☆☆']
hit_star = Counter(r['star'] for r in shot)
MST = r'D:\electrician-2027-0721-review\마스터_1권2권_전체구조_인수인계_2026-07-22_v3.xlsx'
_mw = openpyxl.load_workbook(MST, read_only=True, data_only=True)
_rows = [r for r in _mw['2권 문항마스터'].iter_rows(min_row=2, values_only=True)]
pop_star = Counter(r[6] for r in _rows if r[6])
POPN = sum(pop_star[s] for s in STARS)
hit_round = Counter(re.match(r'(\d+-\d+)회', r['src']).group(1) for r in shot)
TOPR = hit_round.most_common(10)

CSS = r"""
:root{
  --bg:#f4f6f9; --card:#fff; --ink:#151b26; --ink2:#5b6675; --ink3:#8b95a5;
  --line:#e2e6ed; --pink:#c2185b; --pink-bg:#fce4ec;
  --ok:#1f8a4c; --ok-bg:#e6f5ec; --warn:#a8710a; --warn-bg:#fdf3e0;
  --gray:#5b6675; --gray-bg:#eceff4; --shot:#fff;
  --stamp:#d6342c; --navy:#3c4a63; --navy-bg:#eef1f7;
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#11141a; --card:#1a1f28; --ink:#e7eaf0; --ink2:#a3acbd;
    --ink3:#7b8497; --line:#2a3140; --pink:#f06292; --pink-bg:#3a1728;
    --ok:#56c07f; --ok-bg:#14301f; --warn:#d7a63c; --warn-bg:#352a12;
    --gray:#a3acbd; --gray-bg:#252b36; --shot:#fff;
    --stamp:#ff6b62; --navy:#a8b6d1; --navy-bg:#222a38;
  }
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);
  font:16px/1.6 "Pretendard","Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif;
  padding:0 0 64px;overflow-wrap:anywhere;word-break:keep-all}
.wrap{max-width:1180px;margin:0 auto;padding:0 16px}
header{background:var(--card);border-bottom:1px solid var(--line);
  padding:30px 0 26px;margin-bottom:8px}
@media (min-width:860px){header{padding:46px 0 34px}}
.sub{margin:0;color:var(--ink2);font-size:13.5px}

.hero{padding:10px 0 4px}
.hero .kicker{margin:0 0 10px;font-size:13px;font-weight:600;letter-spacing:.04em;
  color:var(--pink);background:var(--pink-bg);display:inline-block;
  padding:6px 13px;border-radius:999px}
.hero h1{font-size:clamp(30px,8.6vw,58px);line-height:1.12;margin:0;
  letter-spacing:-.035em;font-weight:800}
.hero h1 b{color:var(--stamp);font-variant-numeric:tabular-nums}
.hero .lead{margin:12px 0 0;font-size:clamp(14px,3.6vw,17px);color:var(--ink2);line-height:1.5}
.hero .lead b{color:var(--ink);font-weight:600}

.big{display:flex;flex-wrap:wrap;gap:12px;margin:26px 0 4px}
.big div{flex:1 1 168px;border-radius:16px;padding:18px 20px 16px;
  border:1px solid var(--line);background:var(--gray-bg)}
.big .k{font-size:13.5px;font-weight:600;color:var(--ink2);margin:0 0 4px}
.big .v{font-size:clamp(38px,10vw,54px);font-weight:800;letter-spacing:-.04em;
  line-height:1;font-variant-numeric:tabular-nums;margin:0}
.big .v small{font-size:15px;font-weight:600;margin-left:4px;letter-spacing:-.01em}
.big .f{margin:8px 0 0;font-size:12px;color:var(--ink3);line-height:1.4}
.big .p{background:var(--ok-bg);border-color:transparent}
.big .p .k,.big .p .v{color:var(--ok)}
.big .t{background:var(--navy-bg);border-color:transparent}
.big .t .k,.big .t .v{color:var(--navy)}

.stats{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 0;padding:0;list-style:none}
.stats li{background:var(--gray-bg);border:1px solid var(--line);border-radius:10px;
  padding:8px 11px;font-size:12.5px;color:var(--ink2);line-height:1.35}
.stats b{color:var(--ink);font-variant-numeric:tabular-nums;margin-left:5px}
.stats li.wn{background:var(--warn-bg);border-color:transparent}
.stats li.wn b{color:var(--warn)}

nav{display:flex;flex-wrap:wrap;gap:8px;margin:14px 0 0}
nav a{font-size:13.5px;text-decoration:none;color:var(--ink2);background:var(--gray-bg);
  border:1px solid var(--line);border-radius:999px;padding:8px 14px;min-height:38px;
  display:inline-flex;align-items:center}
nav a:hover{color:var(--ink);border-color:var(--ink3)}
.note{margin:14px 0 0;padding:12px 14px;border-left:3px solid var(--pink);
  background:var(--pink-bg);border-radius:0 8px 8px 0;font-size:13.5px;color:var(--ink2)}
.note b{color:var(--ink)}

/* ── 분석 섹션 ───────────────────────────── */
.panel{background:var(--card);border:1px solid var(--line);border-radius:16px;
  padding:18px 16px;margin:0 0 14px}
@media (min-width:860px){.panel{padding:22px 24px}}
.panel h4{margin:0 0 3px;font-size:15px;letter-spacing:-.01em}
.panel .hd{margin:0 0 16px;font-size:12.5px;color:var(--ink3);line-height:1.5}
.rate{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;margin:0 0 12px}
.rate .n{font-size:clamp(34px,9vw,46px);font-weight:800;letter-spacing:-.04em;
  line-height:1;color:var(--ok);font-variant-numeric:tabular-nums}
.rate .d{font-size:13.5px;color:var(--ink2)}
.stack{display:flex;height:34px;border-radius:9px;overflow:hidden;background:var(--gray-bg)}
.stack span{display:flex;align-items:center;justify-content:center;font-size:12px;
  font-weight:700;color:#fff;min-width:0;overflow:hidden;white-space:nowrap}
.s1{background:#1f8a4c}.s2{background:#4a5c7d}.s3{background:#b9c1cf;color:#39404d!important}
@media (prefers-color-scheme:dark){.s3{background:#39404d;color:#c6cddb!important}}
.legend{display:flex;flex-wrap:wrap;gap:12px;margin:12px 0 0;font-size:12.5px;color:var(--ink2)}
.legend i{width:11px;height:11px;border-radius:3px;display:inline-block;
  margin-right:6px;transform:translateY(1px)}

.bars{margin:0;padding:0;list-style:none}
.bars li{display:grid;grid-template-columns:62px 1fr 62px;gap:9px;align-items:center;
  padding:5px 0;font-size:13px}
.bars .lb{color:var(--ink2);letter-spacing:.04em;white-space:nowrap;overflow:hidden;
  text-overflow:ellipsis;font-variant-numeric:tabular-nums}
.bars .tr{display:block;background:var(--gray-bg);border-radius:6px;height:16px;overflow:hidden}
.bars .fi{display:block;height:100%;border-radius:6px;background:var(--pink);min-width:3px}
.bars .fi.alt{background:var(--navy)}
.bars .vl{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink);
  font-weight:700;white-space:nowrap}
.bars .vl small{font-weight:400;color:var(--ink3);margin-left:3px}

.cmp{margin:0;padding:0;list-style:none}
.cmp li{padding:10px 0;border-top:1px solid var(--line)}
.cmp li:first-child{border-top:0;padding-top:0}
.cmp .ch{display:flex;justify-content:space-between;align-items:baseline;gap:10px;
  font-size:13px;margin:0 0 7px;flex-wrap:wrap}
.cmp .ch b{font-size:14px;letter-spacing:.05em}
.cmp .ch em{font-style:normal;color:var(--ink3);font-size:12px;
  font-variant-numeric:tabular-nums}
.cmp .row{display:grid;grid-template-columns:46px 1fr 52px;gap:8px;align-items:center;
  font-size:11.5px;color:var(--ink3);margin:0 0 4px}
.cmp .tr{display:block;background:var(--gray-bg);border-radius:5px;height:12px;overflow:hidden}
.cmp .fi{display:block;height:100%;border-radius:5px;min-width:3px}
.cmp .vl{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink2)}
.readout{margin:14px 0 0;font-size:12.5px;color:var(--ink2);line-height:1.6;
  background:var(--gray-bg);border-radius:10px;padding:11px 13px}
.readout b{color:var(--ink)}

.sec{margin:34px 0 0;scroll-margin-top:10px}
.sec>h2{font-size:19px;margin:0 0 4px;letter-spacing:-.02em;display:flex;
  align-items:baseline;gap:9px;flex-wrap:wrap}
.sec>h2 em{font-style:normal;font-size:13px;font-weight:500;color:var(--ink2)}
.sec>.cap{margin:0 0 4px;font-size:13px;color:var(--ink3)}
.sec>hr{border:0;border-top:2px solid var(--ink);margin:10px 0 18px}
h3.sj{font-size:15px;margin:22px 0 11px;color:var(--ink2);display:flex;align-items:baseline;
  gap:8px;flex-wrap:wrap;scroll-margin-top:10px}
h3.sj:before{content:"";width:4px;height:15px;background:var(--pink);border-radius:2px;
  display:inline-block;transform:translateY(2px)}
h3.sj span{font-size:12.5px;font-weight:400;color:var(--ink3)}

.grid{display:grid;grid-template-columns:1fr;gap:14px}
.grid.two{gap:14px}
@media (min-width:860px){.grid.two{grid-template-columns:1fr 1fr}}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;
  padding:14px;display:flex;flex-direction:column;gap:10px}

/* ── 복원 문항 | 교재 지면 2단 ─────────────── */
.pair{display:grid;grid-template-columns:1fr;gap:14px;align-items:start}
@media (min-width:760px){.pair{grid-template-columns:1fr 1fr;gap:18px}}
.pane{min-width:0}
@media (min-width:760px){.pane.exam{position:sticky;top:14px}}
.pane>.lb{display:flex;align-items:center;gap:7px;flex-wrap:wrap;
  margin:0 0 9px;padding-bottom:7px;border-bottom:1px solid var(--line);
  font-size:12px;font-weight:700;letter-spacing:.02em;color:var(--ink2)}
.pane>.lb .dot{width:7px;height:7px;border-radius:50%;background:var(--pink);flex:none}
.pane.book>.lb .dot{background:var(--navy)}
.qbox{font-size:14.5px;line-height:1.65}
.qbox .qn{display:inline-flex;align-items:center;justify-content:center;
  min-width:26px;height:24px;padding:0 6px;border-radius:7px;background:var(--pink);
  color:#fff;font-size:13px;font-weight:800;margin:0 8px 0 0;
  font-variant-numeric:tabular-nums;vertical-align:2px}
.qbox .stem{display:inline}
.qbox ol{margin:11px 0 0;padding:0;list-style:none;counter-reset:o}
.qbox ol li{counter-increment:o;position:relative;padding:3px 0 3px 25px;
  font-size:14px;line-height:1.6}
.qbox ol li:before{content:counter(o);position:absolute;left:0;top:4px;
  width:17px;height:17px;border-radius:50%;border:1px solid var(--ink3);
  color:var(--ink3);font-size:10.5px;display:flex;align-items:center;
  justify-content:center;font-variant-numeric:tabular-nums}
.qbox ol li.on{font-weight:700;color:var(--ok)}
.qbox ol li.on:before{background:var(--ok);border-color:var(--ok);color:#fff}
.qbox .ans{margin:11px 0 0;font-size:13px;color:var(--ink2)}
.qbox .ans b{color:var(--ok);font-size:14px}
.qbox .fig{margin:10px 0 0;overflow-x:auto}
.qbox .box{border:1px solid var(--ink3);border-radius:7px;padding:9px 12px;
  margin:10px 0 0;font-size:13.5px;line-height:1.7}
.qbox img{max-width:100%;height:auto}
@media (prefers-color-scheme:dark){.qbox img{filter:brightness(.93)}}
.katex{font-size:1em}
.katex-display{margin:.4em 0;overflow-x:auto;overflow-y:hidden}
.card.hit{border-color:color-mix(in srgb,var(--stamp) 35%,var(--line))}
.top{display:flex;gap:9px;align-items:flex-start}
.num{flex:none;width:28px;height:28px;border-radius:8px;background:var(--gray-bg);
  color:var(--ink2);font-size:13px;font-weight:700;display:flex;align-items:center;
  justify-content:center;font-variant-numeric:tabular-nums}
.ttl{font-size:15px;font-weight:700;line-height:1.45;margin:2px 0 0;letter-spacing:-.01em}
.meta{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12.5px}
.tag{border-radius:7px;padding:4px 9px;line-height:1.3;white-space:nowrap}
.tag.src{background:var(--gray-bg);color:var(--ink2);font-variant-numeric:tabular-nums}
.tag.star{background:var(--pink-bg);color:var(--pink);letter-spacing:.06em}
.badge{border:1px solid var(--navy);color:var(--navy);background:transparent;
  border-radius:6px;padding:3px 8px;font-size:11.5px;font-weight:600;white-space:nowrap}
.unit{font-size:12.5px;color:var(--ink3);margin:0}
.hint{font-size:11.5px;color:var(--ink3);margin:-4px 0 0}

.shot{position:relative;background:var(--shot);border:1px solid var(--line);
  border-radius:10px;padding:6px}
.shot img{display:block;width:100%;height:auto;border-radius:5px}
img{max-width:100%;height:auto}
.stamp{position:absolute;top:-13px;right:-5px;z-index:2;
  border:3px double var(--stamp);color:var(--stamp);border-radius:6px;
  padding:5px 11px;font-size:15px;font-weight:800;letter-spacing:.08em;
  line-height:1.2;white-space:nowrap;background:transparent;
  transform:rotate(-14deg);opacity:.9;mix-blend-mode:multiply;
  animation:stamp .45s ease-out both}
@media (prefers-color-scheme:dark){.stamp{mix-blend-mode:normal}}
@keyframes stamp{
  0%{transform:rotate(-14deg) scale(1.6);opacity:0}
  60%{opacity:.95}
  100%{transform:rotate(-14deg) scale(1);opacity:.9}}
@media (prefers-reduced-motion:reduce){
  .stamp{animation:none;transform:rotate(-14deg) scale(1);opacity:.9}}

.chk{font-size:12px;color:var(--ink3);margin:0;display:flex;gap:10px;flex-wrap:wrap}
.chk i{font-style:normal;color:var(--ok)}
.card.miss{border-style:dashed}
.nobook{border:1px dashed var(--line);border-radius:10px;padding:22px 16px;
  display:flex;flex-direction:column;align-items:center;gap:6px;text-align:center;
  background:var(--gray-bg)}
.nobook b{font-size:14px;color:var(--ink2)}
.nobook span{font-size:12.5px;color:var(--ink3);line-height:1.5}
.nobook .src2{font-variant-numeric:tabular-nums}
footer{margin:40px 0 0;padding-top:16px;border-top:1px solid var(--line);
  font-size:12.5px;color:var(--ink3);line-height:1.75}
footer code{background:var(--gray-bg);padding:2px 6px;border-radius:5px;font-size:11.5px}
"""

_delay = [0]


def qpane(r):
    """왼쪽 — 2026년 제4회 복원 문항 전문."""
    o = ['<div class="pane exam">'
         '<p class="lb"><span class="dot"></span>2026년 제4회 복원 문항</p>'
         '<div class="qbox"><span class="qn">%02d</span>'
         '<span class="stem">%s</span>' % (r['n'], r['q_stem'])]
    if r['q_opts']:
        o.append('<ol>')
        for k, op in enumerate(r['q_opts'], 1):
            o.append('<li%s>%s</li>' % (' class="on"' if k == r['q_ans'] else '', op))
        o.append('</ol>')
    if r['q_ans']:
        o.append('<p class="ans">정답 <b>%s</b></p>' % '①②③④'[r['q_ans'] - 1])
    o.append('</div></div>')
    return ''.join(o)


def bpane(r, hit):
    """오른쪽 — 교재 2권 지면 크롭."""
    _delay[0] += 1
    d = min(_delay[0], 12) * 0.06
    o = ['<div class="pane book">'
         '<p class="lb"><span class="dot"></span>우리 교재 · 2권 %s쪽</p>' % r['page']]
    o.append('<div class="meta">'
             '<span class="tag src">2권 %s쪽 · %s</span>' % (r['page'], H.escape(r['src'])))
    if r['star']:
        o.append('<span class="tag star">%s</span>' % H.escape(r['star']))
    if not hit:
        o.append('<span class="badge">유형 일치</span>')
    o.append('</div>')
    if not hit:
        o.append('<p class="hint">같은 단원에서 출제</p>')
    if r['unit']:
        o.append('<p class="unit">관련이론 %s</p>' % H.escape(r['unit']))
    if r['img']:
        o.append('<div class="shot">')
        if hit:
            o.append('<span class="stamp" style="animation-delay:%.2fs">완벽적중</span>' % d)
        o.append('<img src="%s/%s" alt="%s 교재 2권 %s쪽 지면" loading="lazy"></div>'
                 % (RELD, H.escape(r['img']), H.escape(r['src']), r['page']))
    chk = ['<span><i>&#10003;</i> 문항번호 %s</span>' % H.escape(str(r.get('head', '')))]
    if r.get('unit_ok'):
        chk.append('<span><i>&#10003;</i> 소단원 대조</span>')
    if r.get('segs', 1) > 1:
        chk.append('<span>단 넘김 %d컷 이어붙임</span>' % r['segs'])
    if r['status'] == 'page':
        chk.append('<span>쪽 전체</span>')
    if r['note']:
        chk.append('<span>%s</span>' % H.escape(r['note']))
    o.append('<p class="chk">%s</p></div>' % ''.join(chk))
    return ''.join(o)


def card(r, hit):
    return ('<article class="card%s">'
            '<div class="top"><span class="num">%d</span><h3 class="ttl">%s</h3></div>'
            '<div class="pair">%s%s</div></article>'
            % (' hit' if hit else '', r['n'], H.escape(r['title']),
               qpane(r), bpane(r, hit)))


def section(sid, title, sup, cap, rows, hit):
    o = ['<section class="sec" id="%s"><h2>%s <em>%s</em></h2>' % (sid, title, sup)]
    if cap:
        o.append('<p class="cap">%s</p>' % cap)
    o.append('<hr>')
    for s in SUBJS:
        sub = [r for r in rows if r['sn'] == s]
        if not sub:
            continue
        o.append('<h3 class="sj" id="%s-%s">%s <span>%d문항</span></h3>'
                 % (sid, SID[s], s, len(sub)))
        o.append('<div class="grid">')
        o.extend(card(r, hit) for r in sub)
        o.append('</div>')
    o.append('</section>')
    return '\n'.join(o)


# ── 분석 섹션 ─────────────────────────────────────────
def pct(a, b):
    return a / b * 100 if b else 0


an = ['<section class="sec" id="stat"><h2>우리 교재 적중 현황</h2>'
      '<p class="cap">2026년 제4회 필기 60문항을 교재 2권 %s문항과 대조한 결과다.</p><hr>' % format(POPN, ',')]

# 1) 적중 현황
an.append('<div class="panel"><h4>적중률</h4>'
          '<p class="hd">복원 60문항 가운데 교재 2권에 같은 소재가 실려 있던 문항</p>'
          '<div class="rate"><span class="n">%.1f%%</span>'
          '<span class="d">60문항 중 <b>%d문항</b> 적중 · 미수록 %d문항</span></div>'
          % (RATE, HITN, len(MISS)))
an.append('<div class="stack">'
          '<span class="s1" style="width:%.4f%%">%d</span>'
          '<span class="s2" style="width:%.4f%%">%d</span>'
          '<span class="s3" style="width:%.4f%%">%d</span></div>'
          % (pct(len(A), TOTN), len(A), pct(len(W), TOTN), len(W),
             pct(len(MISS), TOTN), len(MISS)))
an.append('<div class="legend">'
          '<span><i class="s1"></i>완벽적중 %d</span>'
          '<span><i class="s2"></i>유형 일치 %d</span>'
          '<span><i class="s3"></i>교재 미수록 %d</span></div></div>'
          % (len(A), len(W), len(MISS)))

# 2) 별점 분포
mx = max(hit_star[s] for s in STARS)
an.append('<div class="panel"><h4>적중 %d문항의 빈출도(별점) 분포</h4>'
          '<p class="hd">별점은 교재 2권 문항마스터 시트의 별표 열 값이다.</p><ul class="bars">'
          % HITN)
for s in STARS:
    an.append('<li><span class="lb">%s</span><span class="tr">'
              '<span class="fi" style="width:%.2f%%"></span></span>'
              '<span class="vl">%d<small>문항</small></span></li>'
              % (s, pct(hit_star[s], mx), hit_star[s]))
an.append('</ul>')
an.append('<p class="readout"><b>★★★ 하나로 적중분의 절반(%d/%d, %.0f%%)을 덮었다.</b> '
          '★★★와 ★★☆를 합치면 %d문항, 적중분의 %.0f%%다.</p></div>'
          % (hit_star['★★★'], HITN, pct(hit_star['★★★'], HITN),
             hit_star['★★★'] + hit_star['★★☆'],
             pct(hit_star['★★★'] + hit_star['★★☆'], HITN)))

# 2-b) 모집단 대비
an.append('<div class="panel"><h4>2권 전체 %s문항 대비 구성비</h4>'
          '<p class="hd">교재가 매긴 별점 구성비와 이번 적중분의 별점 구성비를 나란히 놓았다.</p>'
          '<ul class="cmp">' % format(POPN, ','))
for s in STARS:
    ph, pp = pct(hit_star[s], HITN), pct(pop_star[s], POPN)
    ratio = ph / pp if pp else 0
    an.append('<li><p class="ch"><b>%s</b><em>집중도 %.2f배</em></p>'
              '<div class="row"><span>적중분</span><span class="tr">'
              '<span class="fi" style="width:%.2f%%;background:var(--pink)"></span></span>'
              '<span class="vl">%.1f%%</span></div>'
              '<div class="row"><span>2권 전체</span><span class="tr">'
              '<span class="fi" style="width:%.2f%%;background:var(--navy)"></span></span>'
              '<span class="vl">%.1f%%</span></div></li>'
              % (s, ratio, ph, ph, pp, pp))
an.append('</ul><p class="readout">적중분의 별점 구성비는 2권 전체 구성비와 거의 같다'
          '(집중도 %.2f~%.2f배). <b>별점이 높은 문항이 특별히 더 많이 나왔다기보다, '
          '교재가 매긴 비중 그대로 출제됐다</b>고 읽는 편이 정확하다.</p></div>'
          % (min(pct(hit_star[s], HITN) / pct(pop_star[s], POPN) for s in STARS),
             max(pct(hit_star[s], HITN) / pct(pop_star[s], POPN) for s in STARS)))

# 3) 회차 분포
mxr = TOPR[0][1]
an.append('<div class="panel"><h4>재출제 회차 분포 <span></span></h4>'
          '<p class="hd">적중 %d문항이 2권의 어느 회차 문항과 겹쳤는지 (상위 %d개)</p>'
          '<ul class="bars">' % (HITN, len(TOPR)))
for k, v in TOPR:
    an.append('<li><span class="lb">%s회</span><span class="tr">'
              '<span class="fi alt" style="width:%.2f%%"></span></span>'
              '<span class="vl">%d<small>문항</small></span></li>'
              % (k, pct(v, mxr), v))
an.append('</ul><p class="readout"><b>%s회가 %d문항으로 가장 많았다.</b> '
          '적중분이 걸친 회차는 모두 %d개 회차다.</p></div>'
          % (TOPR[0][0], TOPR[0][1], len(hit_round)))
an.append('</section>')

body = [''.join(an),
        section('perfect', '완벽적중', '완전 일치 %d문항' % len(A),
                '시험 발문과 보기가 교재 2권 원문 그대로다.', A, True),
        section('type', '유형 일치', '%d문항' % len(W),
                '같은 단원에서 출제', W, False)]

miss = MISS
body.append('<section class="sec" id="miss"><h2>교재 미수록 <em>%d문항</em></h2>'
            '<p class="cap">교재 2권 기출에 없어 크롭 대상이 아니다.</p><hr><div class="grid">'
            % len(miss))
for r in miss:
    body.append('<article class="card miss"><div class="top"><span class="num">%d</span>'
                '<h3 class="ttl">%s</h3></div><div class="pair">%s'
                '<div class="pane book"><p class="lb"><span class="dot"></span>우리 교재</p>'
                '<div class="nobook"><b>교재 미수록</b><span>%s</span>'
                '<span class="src2">%s · %s</span></div></div></div></article>'
                % (r['n'], H.escape(r['title']), qpane(r),
                   H.escape(r['note']), H.escape(r['sn']), H.escape(r['src'])))
body.append('</div></section>')

nimg = sum(1 for r in R if r['img'])
doc = """<!DOCTYPE html><html lang="ko"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>2026년 4회 기출복원 &times; 교재 2권 지면</title>
<link rel="stylesheet" href="CBT_2019_1회/katex/katex.min.css">
<script defer src="CBT_2019_1회/katex/katex.min.js"></script>
<script defer src="CBT_2019_1회/katex/auto-render.min.js"></script>
<style>%s</style></head><body>
<header><div class="wrap">
<div class="hero">
<p class="kicker">2027 대비 독끝 전기기능사 필기 2권</p>
<h1>총 <b>%d문항</b> 적중!</h1>
<p class="lead"><b>2026년 제4회 전기기능사 필기 60문항</b> 중 — 복원 문항을 교재 2권 지면과 하나씩 대조했다.</p>
</div>
<div class="big">
<div class="p"><p class="k">완벽적중</p><p class="v">%d<small>문항</small></p></div>
<div class="t"><p class="k">유형 일치</p><p class="v">%d<small>문항</small></p>
<p class="f">같은 단원에서 출제</p></div>
</div>
<ul class="stats">
<li>복원 문항 전체<b>%d</b></li><li>크롭 성공<b>%d</b></li>
<li class="%s">쪽 전체 대체<b>%d</b></li><li class="%s">크롭 실패<b>%d</b></li>
<li>교재 미수록<b>%d</b></li><li>이미지<b>%d컷</b></li>
</ul>
<nav><a href="#stat">적중 현황</a><a href="#perfect">완벽적중</a>
<a href="#type">유형 일치</a><a href="#miss">교재 미수록</a></nav>
<p class="note"><b>읽는 법.</b> <b>완벽적중</b>은 시험 문항이 교재 지면 문항 그대로 나온 것이다.
<b>유형 일치</b>는 시험 문항과 <b>같은 단원에서 출제</b>된 교재 문항으로, 출처의 회차·번호는 집필 때 고른 대표 출처다.
별표는 교재에 매긴 빈출도다.</p>
</div></header>
<div class="wrap">
%s
<footer>
크롭 원본 <code>D:\\CBT 모의고사\\기출\\[2027 전기기능사 필기(2권)]-최종_20260728.pdf</code> · 180dpi ·
문항 번호를 찾아 다음 문항 직전까지 단(column) 단위로 잘랐고, 단을 넘어가는 문항은 두 컷을 이어 붙였다.<br>
쪽 매핑 <b>인쇄쪽 = PDF 페이지(1-based)</b> — 536쪽 러닝헤더 전수 대조, 불일치 0.<br>
검증 크롭 첫 줄 문항번호 %d/%d 일치 · 지면 관련이론 소단원 &harr; 마스터 시트 소단원 %d/%d 일치.<br>
생성 %s
</footer>
</div>
<script>document.addEventListener("DOMContentLoaded",function(){
if(window.renderMathInElement){renderMathInElement(document.body,{delimiters:[
{left:"$$",right:"$$",display:true},{left:"$",right:"$",display:false}],
throwOnError:false});}});</script>
</body></html>""" % (
    CSS, len(A) + len(W), len(A), len(W),
    len(R), len(ok), 'wn' if pgs else '', len(pgs), 'wn' if fls else '', len(fls), len(sks), nimg,
    '\n'.join(body),
    sum(1 for r in ok if r.get('head')), len(ok),
    sum(1 for r in ok if r.get('unit_ok') is True), len(ok),
    datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))

open(OUTH, 'w', encoding='utf-8').write(doc)
png = [f for f in os.listdir(OUTD) if f.lower().endswith('.png')]
print('저장:', OUTH)
print('완벽적중 %d · 유형 일치 %d · 교재 미수록 %d · 전체 %d' % (len(A), len(W), len(sks) + len(fls), len(R)))
print('img 태그 %d / PNG 파일 %d / 도장 %d' % (doc.count('<img '), len(png), doc.count('class="stamp"')))

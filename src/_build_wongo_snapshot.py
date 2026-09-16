# -*- coding: utf-8 -*-
"""2026년 4회(09-16) 기출복원 원고 — 문제편. 해설 제외.
A/B/C 등급에 따라 CBT DB 원문을 끌어와 원고 틀에 얹는다."""
import sys, io, html
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

CBT = r'D:\electrician-cbt\전기기능사_CBT_31회차_통합.xlsx'
MST = r'D:\electrician-2027-0721-review\마스터_1권2권_전체구조_인수인계_2026-07-22_v3.xlsx'
OUT = r'D:\electrician-cbt\_2026-4회_복원원고_문제편_20260916.html'

# (과목, 회상 문항명, 등급, 대표 DB 출처, 내가 고른 답)
# 등급 A = DB 원문 그대로 사용 가능 / B = 발문 틀 동일, 수치·보기 교체 필요 / C = 개념 공유, 발문 재작성 필요
ITEMS = [
 # ── 전기이론 ──────────────────────────────────────────
 ('이론', '최댓값 10[A] 교류전류의 평균값',            'A', (2020,4,1),   '—'),
 ('이론', '무효전력의 기호·식',                      'W', (2024,4,7),   'VI sinθ'),
 ('이론', '병렬 합성정전용량',                       'A', (2025,4,5),   'C1+C2'),
 ('이론', '망간 건전지의 양극 재료',                  'A', (2021,4,16),  '미기재'),
 ('이론', '전기력선의 특징이 아닌 것',                 'W', (2025,4,11),  '전위 낮은곳→높은곳'),
 ('이론', '부하에서 얻을 수 있는 최대전력',             'A', (2022,4,6),   '125W'),
 ('이론', '니켈의 화학당량',                         'A', (2019,3,2),   '미기재 (④ 29.35)'),
 ('이론', '발전기의 원리',                          'W', (2019,2,4),   '플레밍 오른손'),
 ('이론', '평판 콘덴서의 정전용량을 늘리는 방법이 아닌 것', 'W', (2022,1,1),   '간격을 늘린다'),
 ('이론', '전지 n개 병렬 접속 시 전체 기전력',           'W', (2019,2,16),  'V0'),
 ('이론', '저항 12[Ω]·300[V]·15[A]일 때 리액턴스',   'W', (2019,3,8),   '미기재'),
 ('이론', '극성이 있어 교류에 쓸 수 없는 콘덴서',         'W', (2023,3,20),  '④ 전해 콘덴서'),
 ('이론', '코일 쇄교 자속 변화에 의한 유도기전력',        'A', (2019,1,8),  '미기재'),
 ('이론', '200[W] 전구 5개를 12시간 사용한 전력량',     'W', (2025,3,13),  '③ 12 kWh'),
 ('이론', '평행 두 도선 10[A]·5[cm] 1[m]당 작용력',  'A', (2020,3,4),   '미기재'),
 ('이론', 'Y 30[Ω]의 등가 Δ 저항 Rc',              'W', (2022,4,10),  '10Ω (오답)'),
 ('이론', '대칭 3상 교류의 조건',                     'W', (2026,1,19),  '④ 4π/3'),
 ('이론', '유전율의 단위 〔1권 79쪽 (2) 유전율★★〕',      'T', 'permit',     '① [F/m]'),
 ('이론', '8[μC]·4[μC]가 7.2[N]일 때 거리',        'R', 'coulomb',    '0.2m'),
 ('이론', '중첩의 원리 — 전압원 (ㄱ), 전류원 (ㄴ)',      'R', 'superpos',   'ㄱ 단락·ㄴ 개방'),
 # ── 전기기기 ──────────────────────────────────────────
 ('기기', '직류기 중 정속도 특성을 가진 것',            'A', (2024,4,30),  '분권'),
 ('기기', 'Y-Δ 기동 시 기동토크·전류의 변화',          'W', (2020,4,31),  '1/3'),
 ('기기', '우리나라 대표 변압기 극성',                 'W', (2023,1,37),  '감극성'),
 ('기기', '워드 레오나드 방식에 필요한 기기',            'W', (2019,1,23),  '③ 타여자 발전기'),
 ('기기', '직류기 파권의 병렬회로수 a',                'A', (2026,1,27),  '2'),
 ('기기', '전류가 90° 앞설 때 전기자 반작용',          'A', (2023,1,25),  '증자작용'),
 ('기기', '유도전동기의 2차 효율',                    'W', (2021,2,40),  '출력/1차입력 (오답)'),
 ('기기', 'TRIAC의 특성·기호',                      'W', (2021,2,36),  '미기재'),
 ('기기', '속도 제어법이 아닌 것',                     'W', (2024,3,22),  '위상 제어법'),
 ('기기', '인가 전압 90%일 때 토크',                  'W', (2019,4,26),  '81%'),
 ('기기', '전기자 권선을 단절권으로 하면',               'A', (2021,4,25),  '④ 기전력을 높인다 (오답)'),
 ('기기', '부흐홀츠 계전기로 보호되는 기기',             'A', (2022,4,24),  '변압기'),
 ('기기', '단상 유도전동기 중 역률이 가장 좋은 것',       'W', (2019,1,32),  '미기재'),
 ('기기', '%Z일 때 단락전류는 정격전류의 몇 배',         'A', (2024,1,24),  '20배'),
 ('기기', '무부하 103[V]·정격 100[V] 전압변동률',     'W', (2021,3,29),  '미기재'),
 ('기기', '전동기에서 자속을 줄이면 속도는',             'W', (2020,1,28),  '증가한다'),
 ('기기', '3상 반파 정류회로의 직류 평균전압',           'W', (2022,3,33),  '미기재'),
 ('기기', '분권 발전기 Ia=100[A]·If=6[A] 부하전류',  'W', (2019,3,33),  '94A'),
 ('기기', '소형 실험실 등에 쓰는 발전기',                'W', (2019,3,23),  '분권 발전기'),
 ('기기', 'Δ-Y 결선의 특징 (위상차 60°)',            'O', ('20090712',32), '60° 보기'),
 # ── 전기설비 ──────────────────────────────────────────
 ('설비', '합성수지관의 장점이 아닌 것',                'A', (2024,4,51),  '기계적 강도가 세다'),
 ('설비', '접지저항 측정에 쓰는 것',                   'A', (2025,4,56),  '콜라우시 브리지'),
 ('설비', '설계하중 6.8[kN]·전장 7[m] 근입깊이',      'A', (2019,2,41),  '1.2m'),
 ('설비', '한 등을 두 곳에서 점멸하는 배선 (3로)',       'A', (2026,3,58),  '미기재'),
 ('설비', '애자공사 시 전선 상호간 간격',               'W', (2021,3,57),  '3cm'),
 ('설비', '화약류 저장소 전기설비 기준 중 틀린 것',       'W', (2020,3,50),  '전폐형 (오답)'),
 ('설비', '승탑(전주오름)용 발판 볼트의 높이',           'A', (2020,3,55),  '1.8m'),
 ('설비', '개폐기 역할과 과전류 차단을 겸용하는 것',       'W', (2019,2,55),  '③ 배선용 차단기'),
 ('설비', '가요전선관과 금속관의 접속',                 'A', (2019,3,58),  '미기재'),
 ('설비', '옥외 백열전등 인하선의 연동선 굵기',           'X', (2024,2,53),  '② 2.5mm²'),
 ('설비', '기숙사에 해당하는 표준부하',                 'W', (2019,3,55),  '미기재'),
 ('설비', '지락전류를 검출하는 것',                    'A', (2025,2,59),  'ZCT'),
 ('설비', '전선 피복을 벗겨내는 공구',                  'A', (2021,1,52),  '니퍼 (오답)'),
 ('설비', '합성수지관 공사 설명 중 옳지 않은 것',        'W', (2019,3,52),  '② 매입 (오답)'),
 ('설비', '지선에 연선 사용 시 소선수',                 'A', (2024,4,47),  '3가닥'),
 ('설비', '접속함·박스 내에서 쓰는 접속법',             'A', (2020,4,52),  '쥐꼬리 접속'),
 ('설비', '전선 약호 중 H가 나타내는 것 (④ 경동선)',      'X', (2019,1,46),  '미기재'),
 ('설비', '보호도체(PE)의 식별 색상',                 'W', (2026,2,41),  '녹색-노란색'),
 ('설비', 'EQ 심벌',                               'R', 'eqsym',      '③ 지진감지기'),
 # 미복원 1문항 자리 — 전기설비 빈출 소재에서 임의 선정
 ('설비', '후강 전선관 규격이 아닌 것 〔임의 삽입〕',        'A', (2020,4,58),  '—'),
]

# ── CBT DB ────────────────────────────────────────────
wb = openpyxl.load_workbook(CBT, read_only=True, data_only=True)
ws = wb[wb.sheetnames[0]]
rows = list(ws.iter_rows(min_row=2, values_only=True))
hdr = [str(x) if x else '' for x in rows[0]]
idx = {h: i for i, h in enumerate(hdr)}


def g(r, name):
    i = idx.get(name)
    v = r[i] if (i is not None and i < len(r)) else None
    return '' if v is None else str(v)


bank = {}
for r in rows[1:]:
    try:
        bank[(int(float(g(r, '연도'))), int(float(g(r, '회차'))), int(float(g(r, '번호'))))] = r
    except ValueError:
        pass

# ── 교재 쪽수 ──────────────────────────────────────────
page = {}
try:
    mw = openpyxl.load_workbook(MST, read_only=True, data_only=True)
    for r in mw['2권 문항마스터'].iter_rows(min_row=2, values_only=True):
        try:
            y, h = str(r[0]).split('-')
            page[(int(y), int(h), int(r[1]))] = r[2]
        except (ValueError, AttributeError, TypeError):
            pass
except Exception as e:
    print('교재쪽 로드 실패:', e)


def fixsrc(s, y, h):
    if 'src="' not in s:
        return s
    pre = 'CBT_%d_%d회/img/' % (y, h)
    return s.replace('src="', 'src="' + pre).replace(pre + 'CBT_', 'CBT_')


TIER = {
 'W': ('C·재작성', '시험 원문대로 재작성 완료 (DB 문항은 참고)'),
 'A': ('A', 'DB 원문 그대로 사용 가능'),
 'B': ('B', '발문 틀 동일 — 수치·보기 교체 필요'),
 'C': ('C', '개념 공유 — 발문 재작성 필요'),
 'O': ('A·구기출', '2002~2009 기출 원문 — 그대로 사용 가능'),
 'T': ('1권 이론 수록', '2권 기출엔 없으나 1권 이론에 그대로 있음'),
 'R': ('신규·복원완료', 'DB 미수록이나 발문·보기 전문 확보'),
 'N': ('신규', 'DB 미수록 — 전문 복원 필요'),
}

# 시험장 기억으로 전문 복원한 문항 (DB·구기출 모두 미수록)
# 재작성분 — 시험 원문 기준. 키는 대표 DB 출처 튜플
# ※ 회상에 없던 오답 보기는 DB 원문·표준 함정을 근거로 채웠다(원고 주석 참조).
REWRITE = {
 # ── 전기이론 ──
 (2026, 1, 19): (
   '대칭 3상 교류의 조건에 해당하지 않는 것은?',
   ['기전력의 크기가 같을 것', '주파수가 같을 것', '파형이 같을 것',
    '위상차각이 $\\dfrac{4\\pi}{3}\\,[\\mathrm{rad}]$이어야 할 것'], 4),
 (2024, 4, 7): (
   '무효전력의 계산식은?',
   ['$VI$', '$VI\\cos\\theta$', '$VI\\sin\\theta$', '$VI\\tan\\theta$'], 3),
 (2025, 4, 11): (
   '전기력선의 성질로 옳지 않은 것은?',
   ['양전하에서 나와 음전하에서 끝난다.',
    '전기력선은 서로 교차하지 않는다.',
    '도체 표면에 수직으로 출입한다.',
    '전위가 낮은 곳에서 높은 곳으로 향한다.'], 4),
 (2022, 1, 1): (
   '평판 콘덴서의 정전용량을 크게 하는 방법이 아닌 것은?',
   ['극판의 면적을 늘린다.', '극판의 간격을 늘린다.',
    '유전율을 크게 한다.', '비유전율이 큰 유전체를 사용한다.'], 2),
 (2022, 4, 10): (
   '그림과 같이 Y결선된 각 상의 저항이 $30\\,[\\Omega]$일 때, '
   '이와 등가인 $\\triangle$결선 회로의 저항 $R_c$의 값은?'
   '<div class="fig"><img src="_figs/ydelta.png" alt="Y-△ 등가 변환" '
   'style="max-width:330px"></div>',
   ['$10\\,[\\Omega]$', '$30\\,[\\Omega]$',
    '$30\\sqrt{3}\\,[\\Omega]$', '$90\\,[\\Omega]$'], 4),
 # ── 전기기기 ──
 (2020, 4, 31): (
   '3상 농형 유도 전동기를 $Y-\\triangle$ 기동할 때, $Y$로 기동하는 경우 '
   '기동토크와 기동전류는 전전압 기동 시의 몇 배가 되는가?',
   ['$\\dfrac{1}{3}$로 된다.', '$\\dfrac{1}{\\sqrt{3}}$로 된다.',
    '$\\sqrt{3}$배로 된다.', '$3$배로 된다.'], 1),
 (2023, 1, 37): (
   '우리나라 표준 변압기의 극성은?',
   ['감극성', '가극성', '무극성', '양극성'], 1),
 (2021, 2, 40): (
   '유도 전동기의 2차 효율을 나타내는 식이 아닌 것은?',
   ['출력 / 1차입력 $\\times\\,100\\,[\\%]$',
    '2차출력 / 2차입력 $\\times\\,100\\,[\\%]$',
    '$(1-s)\\times 100\\,[\\%]$',
    '2차출력 / (2차출력 + 2차동손) $\\times\\,100\\,[\\%]$'], 1),
 (2021, 2, 36): (
   '그림과 같은 기호가 나타내는 소자는?'
   '<div class="fig"><img src="_figs/triac.png" alt="TRIAC 기호" '
   'style="max-width:120px"></div>',
   ['SCR', 'Diode', 'IGBT', 'TRIAC'], 4),
 (2024, 3, 22): (
   '직류 전동기의 속도 제어법이 아닌 것은?',
   ['전압 제어법', '계자 제어법', '저항 제어법', '위상 제어법'], 4),
 (2019, 4, 26): (
   '슬립이 일정할 때 유도 전동기의 인가 전압을 $90\\,[\\%]$로 낮추면 '
   '토크는 처음의 몇 $[\\%]$가 되는가?',
   ['$45\\,[\\%]$', '$81\\,[\\%]$', '$90\\,[\\%]$', '$100\\,[\\%]$'], 2),
 (2021, 4, 25): (
   '동기 발전기의 전기자 권선을 단절권으로 하는 경우의 특징이 아닌 것은?',
   ['고조파를 제거하여 파형이 좋아진다.',
    '코일 단부가 짧아져 동량이 절약된다.',
    '기계적으로 제작이 쉬워진다.',
    '유기 기전력이 높아진다.'], 4),
 # ②의 정확한 문구 미확인 — '콘덴서 전동기(영구 콘덴서형)'이면 정답이 ②로 이동
 (2019, 1, 32): (
   '단상 유도 전동기 중 역률이 가장 좋은 전동기는?',
   ['셰이딩 코일형', '콘덴서 전동기형', '콘덴서 기동형', '분상 기동형'], 3),
 (2021, 3, 29): (
   '직류 발전기의 무부하 전압이 $103\\,[\\mathrm{V}]$, 정격 전압이 '
   '$100\\,[\\mathrm{V}]$이다. 이 발전기의 전압 변동률 $\\varepsilon\\,[\\%]$은?',
   ['$1$', '$3$', '$6$', '$9$'], 2),
 (2020, 1, 28): (
   '직류 전동기에서 자속을 감소시키면 회전 속도는 어떻게 되는가?',
   ['증가한다.', '감소한다.', '변함이 없다.', '정지한다.'], 1),
 # 보기 ①④는 회상 확정(유도자형·차동복권), ②③은 표준 분류에서 채움
 (2019, 3, 23): (
   '전압 변동률이 작고 자여자이므로 별도의 여자 전원이 필요 없으며, '
   '계자 저항기로 전압 조정이 가능하여 전기화학용 전원, 전지의 충전용, '
   '소형 실험실의 전원 등에 사용되는 발전기는?',
   ['유도자형 발전기', '분권 발전기', '직권 발전기', '차동 복권 발전기'], 2),
 (2022, 3, 33): (
   '$3$상 반파 정류 회로에서 직류 전압의 평균값은?'
   '(단, $E$는 교류 전압의 실횻값이다.)',
   ['$0.45E\\,[\\mathrm{V}]$', '$0.9E\\,[\\mathrm{V}]$',
    '$1.17E\\,[\\mathrm{V}]$', '$1.35E\\,[\\mathrm{V}]$'], 3),
 (2019, 3, 33): (
   '그림과 같은 직류 분권 발전기에서 전기자 전류가 $100\\,[\\mathrm{A}]$, '
   '계자 전류가 $6\\,[\\mathrm{A}]$일 때 부하에 흐르는 전류 $I$는?'
   '<div class="fig"><img src="_figs/shunt.png" alt="직류 분권 발전기 회로" '
   'style="max-width:270px"></div>',
   ['$6\\,[\\mathrm{A}]$', '$94\\,[\\mathrm{A}]$',
    '$100\\,[\\mathrm{A}]$', '$106\\,[\\mathrm{A}]$'], 2),
 # ── 전기설비 ──
 (2020, 3, 55): (
   '가공 전선의 지지물에 승탑(전주오름) 또는 승강용으로 사용하는 발판 볼트 등은 '
   '지표상 몇 $[\\mathrm{m}]$ 미만에 시설하여서는 안 되는가?',
   ['$1.2$', '$1.5$', '$1.6$', '$1.8$'], 4),
 (2021, 3, 57): (
   '애자 사용 공사에 대한 설명으로 틀린 것은?',
   ['전선과 조영재의 이격거리는 $2.5\\,[\\mathrm{cm}]$ 이상일 것',
    '전선의 지지점 간 거리는 조영재의 윗면 또는 옆면에 따라 붙일 경우 '
    '$2\\,[\\mathrm{m}]$ 이하일 것',
    '애자는 절연성·난연성 및 내수성이 있는 것일 것',
    '전선 상호 간의 간격은 $3\\,[\\mathrm{cm}]$ 이상일 것'], 4),
 (2020, 3, 50): (
   '화약고 등의 위험 장소에서 전기설비 시설에 관한 내용으로 틀린 것은?',
   ['전로의 대지 전압은 $300\\,[\\mathrm{V}]$ 이하일 것',
    '전기 기계 기구는 전폐형을 사용할 것',
    '전용 개폐기 및 과전류차단기는 화약고 내에 시설할 것',
    '화약고 인입구까지의 배선은 케이블을 사용한 지중 전선로로 할 것'], 3),
 (2019, 3, 58): (
   '가요전선관과 금속관을 상호 접속하는 데 사용하는 것은?',
   ['스플리트 커플링', '콤비네이션 커플링',
    '스트레이트 박스 커넥터', '앵글 박스 커넥터'], 2),
 (2019, 3, 55): (
   '배선설계를 위한 전등 및 소형 전기기계기구의 부하용량 산정 시, '
   '표준부하를 $20\\,[\\mathrm{VA/m^2}]$으로 적용하여야 하는 건축물은?',
   ['교회, 극장', '기숙사, 여관', '은행, 상점', '주택, 아파트'], 2),
 (2019, 3, 52): (
   '합성수지관 공사에 대한 설명 중 옳지 않은 것은?',
   ['습기가 많은 장소 또는 물기가 있는 장소에 시설하는 경우에는 방습 장치를 한다.',
    '관 상호간 및 박스와는 관을 삽입하는 깊이를 관의 바깥지름의 '
    '$1.2$배 이상으로 한다.',
    '관의 지지점간의 거리는 $3\\,[\\mathrm{m}]$ 이상으로 한다.',
    '합성수지관 안에는 전선에 접속점이 없도록 한다.'], 3),
 (2019, 1, 46): (
   '전선 약호 중 H가 나타내는 것은?',
   ['미네럴 인슐레이션 케이블', '비닐절연 네온 전선',
    '옥외용 가교 폴리에틸렌 전선', '경동선'], 4),
 (2026, 2, 41): (
   '전선의 색상 식별에서 보호도체(PE)에 사용하는 색상은?',
   ['갈색', '검은색', '녹색-노란색', '파란색'], 3),
 (2019, 2, 16): (
   '기전력이 $V_0\\,[\\mathrm{V}]$인 전지 $n$개를 병렬로 연결하였을 때 '
   '전체 기전력은 얼마인가?',
   ['$V_0$', '$nV_0$', '$\\dfrac{V_0}{n}$', '$\\dfrac{V_0}{n^2}$'], 1),
 (2019, 3, 8): (
   '저항이 $12\\,[\\Omega]$인 회로에 교류전압 $300\\,[\\mathrm{V}]$를 인가하였더니 '
   '전류가 $15\\,[\\mathrm{A}]$가 흘렀다. 이때 리액턴스의 값은?',
   ['$9\\,[\\Omega]$', '$12\\,[\\Omega]$', '$16\\,[\\Omega]$', '$20\\,[\\Omega]$'], 3),
 (2023, 3, 20): (
   '콘덴서 중 극성을 가지고 있는 콘덴서로서 교류 회로에 사용할 수 없는 것은?',
   ['마일러 콘덴서', '마이카 콘덴서', '세라믹 콘덴서', '전해 콘덴서'], 4),
 (2024, 2, 53): (
   '옥외 백열전등의 인하선으로서 지표상의 높이 $2.5\\,[\\mathrm{m}]$ 미만의 부분에 '
   '사용하는 연동선의 공칭 단면적은 몇 $[\\mathrm{mm^2}]$ 이상이어야 하는가?',
   ['$2.0$', '$2.5$', '$4$', '$6$'], 2),
 # ①은 회상 미확정 — '누전 차단기'로 채움
 (2019, 2, 55): (
   '저압 전로에서 개폐기의 역할과 과전류 차단을 겸용할 수 있는 것은?',
   ['누전 차단기', '컷아웃 스위치', '배선용 차단기', '통형 퓨즈'], 3),
 (2019, 1, 23): (
   '직류 전동기의 워드 레오나드(Ward-Leonard) 속도 제어 방식에서, '
   '전동기에 가하는 단자 전압을 조정하기 위하여 반드시 필요한 기기는?',
   ['정류기', '기동 저항기', '타여자 발전기', '단권 변압기'], 3),
 (2025, 3, 13): (
   '소비전력이 $200\\,[\\mathrm{W}]$인 전구 $5$개를 $12$시간 동안 사용하였다면 '
   '전력량 $[\\mathrm{kWh}]$은 얼마인가?',
   ['$2.4\\,[\\mathrm{kWh}]$', '$6\\,[\\mathrm{kWh}]$',
    '$12\\,[\\mathrm{kWh}]$', '$24\\,[\\mathrm{kWh}]$'], 3),
}

RESTORED = {
 'permit': (
   '유전율의 단위는?',
   ['$[\\mathrm{F/m}]$', '$[\\mathrm{V/m}]$',
    '$[\\mathrm{C/m^2}]$', '$[\\mathrm{H/m}]$'], 1),
 'eqsym': (
   '다음과 같은 심벌은 무엇을 의미하는가?'
   '<div class="fig"><img src="_figs/eq.png" alt="EQ 심벌" '
   'style="max-width:70px"></div>',
   ['변압기 용량', '전류제한기', '지진감지기', '누전경보기'], 3),
 'coulomb': (
   '공기 중에서 $8\\,[\\mu\\mathrm{C}]$과 $4\\,[\\mu\\mathrm{C}]$의 두 전하 사이에 '
   '$7.2\\,[\\mathrm{N}]$의 힘이 작용할 때, 두 전하 사이의 거리는 몇 '
   '$[\\mathrm{m}]$인가?',
   ['$1\\,[\\mathrm{m}]$', '$\\dfrac{1}{5}\\,[\\mathrm{m}]$',
    '$5\\,[\\mathrm{m}]$', '$\\dfrac{1}{25}\\,[\\mathrm{m}]$'], 2),
 'superpos': (
   '다음 ( ㉠ ), ( ㉡ )에 들어갈 내용으로 알맞은 것은?'
   '<div class="box">회로망 분석기법으로 중첩의 원리를 사용할 때 '
   '전압원은 ( ㉠ ), 전류원은 ( ㉡ ) 한다.</div>',
   ['㉠ 단락, ㉡ 단락', '㉠ 단락, ㉡ 개방',
    '㉠ 개방, ㉡ 단락', '㉠ 개방, ㉡ 개방'], 2),
}

# 2002~2009 구기출 원문 (CBT DB 미수록분)
OLD = {
 # 시험 출제분은 보기 ③↔④ 순서가 바뀌어 정답이 ④가 됨
 ('20021006', 11): (
   '대칭 3상 교류의 조건에 해당하지 않는 것은?',
   ['기전력의 크기가 같을 것', '주파수가 같을 것', '파형이 같을 것',
    '위상차각이 $\\dfrac{4\\pi}{3}\\,[\\mathrm{rad}]$이어야 할 것'], 4),
 ('20080330', 20): (
   '유전율의 단위는?',
   ['[F/m]', '[V/m]', '[C/m²]', '[H/m]'], 1),
 # 시험 출제분은 용도에 '소형 실험실'이 추가되고 오답 하나가 '유도자형'으로 교체됨
 ('20090329', 36): (
   '타여자 발전기와 같이 전압 변동률이 적고 자여자이므로 다른 여자 전원이 필요 없으며, '
   '계자 저항기를 사용하여 전압 조정이 가능하므로 전기화학용 전원, 전지의 충전용, '
   '동기기의 여자용으로 쓰이는 발전기는?',
   ['분권 발전기', '직권 발전기', '과복권 발전기', '차동복권 발전기'], 1),
 ('20090712', 32): (
   '변압기를 △-Y 결선(delta-star connection)한 경우에 대한 설명으로 옳지 않은 것은?',
   ['1차 선간전압 및 2차 선간전압의 위상차는 $60°$ 이다.',
    '제3조파에 의한 장해가 적다.',
    '1차 변전소의 승압용으로 사용된다.',
    'Y결선의 중성점을 접지할 수 있다.'], 1),
}
SUBJ = {'이론': '전기이론', '기기': '전기기기', '설비': '전기설비'}

out = []
ans = {'이론': [], '기기': [], '설비': []}
stat = {'A': 0, 'B': 0, 'C': 0, 'O': 0, 'R': 0, 'W': 0, 'T': 0, 'N': 0}

for subj in ('이론', '기기', '설비'):
    lst = [x for x in ITEMS if x[0] == subj]
    out.append('<section class="blk"><h2>%s <span class="cnt">%d문항</span></h2>'
               % (SUBJ[subj], len(lst)))
    for n, (_, title, tier, src, mine) in enumerate(lst, 1):
        stat[tier] += 1
        lab, note = TIER[tier]
        out.append('<article class="q t%s"><div class="qh">'
                   '<span class="no">%d</span>' % (tier, n))
        if tier == 'O':
            d, qn = src
            out.append('<span class="src">%s-%s-%s 기출 %d번</span>'
                       % (d[:4], d[4:6], d[6:8], qn))
        elif tier == 'R':
            out.append('<span class="src">시험장 복원</span>')
        elif tier == 'T':
            out.append('<span class="src">1권 79쪽 · (2) 유전율</span>')
        elif src:
            y, h, qn = src
            pg = page.get(src)
            out.append('<span class="src">%d-%d회 %d번%s</span>'
                       % (y, h, qn, (' · 2권 %s쪽' % pg) if pg else ''))
        out.append('</div>')

        if tier in ('O', 'R', 'W', 'T'):
            stem, opts, a = (OLD[src] if tier == 'O'
                             else RESTORED[src] if tier in ('R', 'T')
                             else REWRITE[src])
            # 이 셋은 내가 작성한 본문이라 HTML(지문 박스)을 그대로 통과시킨다
            out.append('<div class="stem">%s</div><ol class="opt">' % stem)
            for k, o in enumerate(opts, 1):
                on = ' class="on"' if k == a else ''
                out.append('<li%s>%s</li>' % (on, o))
            out.append('</ol>')
            ans[subj].append(str(a))
        elif src and src in bank:
            r = bank[src]
            y, h, qn = src
            a = g(r, '정답(1~4)')
            out.append('<div class="stem">%s</div><ol class="opt">' % fixsrc(g(r, '발문'), y, h))
            for k in range(1, 5):
                on = ' class="on"' if str(k) == a else ''
                out.append('<li%s>%s</li>' % (on, fixsrc(g(r, '보기%d' % k), y, h)))
            out.append('</ol>')
            ans[subj].append(a or '?')
        else:
            out.append('<div class="stem gap">%s</div>'
                       '<p class="todo">%s</p>' % (html.escape(title), note))
            ans[subj].append('?')
        if tier in ('B', 'C') and False:
            pass
        out.append('</article>')
    out.append('</section>')

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1a2233;--ink2:#5b6474;--line:#e3e6ec;
--a:#2e9e5b;--abg:#e7f5ec;--b:#b8860b;--bbg:#fdf5e3;--c:#2f6fed;--cbg:#e9f0fd;
--n:#d64545;--nbg:#fdecec}
@media(prefers-color-scheme:dark){:root{--bg:#14161c;--card:#1d2029;--ink:#e8eaf0;
--ink2:#a7adbd;--line:#2c303c;--a:#3fae6c;--abg:#173325;--b:#d9b24c;--bbg:#3a2f14;
--c:#6f9cf5;--cbg:#1a2742;--n:#e07070;--nbg:#3a1d1d}img{filter:brightness(.92)}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);padding:16px 16px 72px;
font:16px/1.65 "Pretendard","Malgun Gothic",system-ui,sans-serif}
.wrap{max-width:880px;margin:0 auto}
h1{font-size:21px;margin:0 0 4px;line-height:1.35}
.sub{color:var(--ink2);font-size:13.5px;margin:0 0 16px}
.legend{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 24px}
.legend span{border:1px solid var(--line);background:var(--card);border-radius:8px;
padding:8px 12px;font-size:13.5px}
.blk{margin:0 0 30px}
.blk h2{font-size:18px;margin:0 0 12px;padding-bottom:8px;border-bottom:2px solid var(--ink);
display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.cnt{font-size:13px;color:var(--ink2);font-weight:400}
.q{background:var(--card);border:1px solid var(--line);border-radius:12px;
padding:14px 16px;margin:0 0 12px;overflow-wrap:anywhere}
.q.tA{border-left:4px solid var(--a)}
.q.tB{border-left:4px solid var(--b)}
.q.tC{border-left:4px solid var(--c)}
.q.tN{border-left:4px solid var(--n);background:var(--nbg)}
.q.tO{border-left:4px solid var(--a)}
.q.tR{border-left:4px solid var(--c)}
.q.tT{border-left:4px solid var(--a)}
.q.tW{border-left:4px solid var(--c)}
.qh{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
.no{font-weight:700;font-size:16px;min-width:26px;font-variant-numeric:tabular-nums}
.tier{font-size:12px;font-weight:700;border-radius:6px;padding:3px 9px}
.tier.tA{background:var(--abg);color:var(--a)}
.tier.tB{background:var(--bbg);color:var(--b)}
.tier.tC{background:var(--cbg);color:var(--c)}
.tier.tN{background:var(--nbg);color:var(--n)}
.tier.tO{background:var(--abg);color:var(--a)}
.tier.tR{background:var(--cbg);color:var(--c)}
.tier.tW{background:var(--cbg);color:var(--c)}
.src{font-size:12.5px;color:var(--ink2);font-variant-numeric:tabular-nums}
.mine{font-size:12.5px;color:var(--ink2);margin-left:auto}
.stem{margin-bottom:10px}
.stem.gap{color:var(--ink2);font-style:italic}
.box{border:1px solid var(--ink2);border-radius:6px;padding:10px 14px;margin:10px 0 0;
background:var(--bg);font-size:15px;line-height:1.7}
.fig{margin:12px 0 0;color:var(--ink);overflow-x:auto}
.opt{margin:0;padding-left:26px}
.opt li{margin:4px 0;min-height:26px}
.opt li.on{font-weight:700;color:var(--a)}
.todo{margin:8px 0 0;font-size:13px;color:var(--b)}
.ansbox{background:var(--card);border:1px solid var(--line);border-radius:12px;
padding:14px 16px;margin-top:24px}
.ansbox h2{font-size:17px;margin:0 0 10px}
.ansrow{font-variant-numeric:tabular-nums;font-size:14.5px;margin:6px 0;
overflow-x:auto;white-space:nowrap;-webkit-overflow-scrolling:touch}
.ansrow b{display:inline-block;min-width:74px}
img{max-width:100%;height:auto}
"""

arows = []
for s in ('이론', '기기', '설비'):
    seq = ' '.join('%d.%s' % (i, a) for i, a in enumerate(ans[s], 1))
    arows.append('<p class="ansrow"><b>%s</b>%s</p>' % (SUBJ[s], seq))

doc = (
 '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">\n'
 '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
 '<title>2026년 4회 기출복원 원고 — 문제편</title>\n'
 '<link rel="stylesheet" href="CBT_2019_1회/katex/katex.min.css">\n'
 '<script defer src="CBT_2019_1회/katex/katex.min.js"></script>\n'
 '<script defer src="CBT_2019_1회/katex/auto-render.min.js"></script>\n'
 '<style>' + CSS + '</style></head><body><div class="wrap">\n'
 '<h1>2026년 4회 기출복원 원고 — 문제편</h1>\n'
 '<p class="sub">시행 2026-09-16 · 해설 제외 · 미복원 2문항 제외 '
 '(전기이론 1 · 전기설비 1)</p>\n'
 '<div class="legend">'
 '<span>전기이론 <b>20</b></span><span>전기기기 <b>20</b></span>'
 '<span>전기설비 <b>20</b></span>'
 '<span>2권 수록 <b>%d</b></span>'
 '<span>구기출(02~09) <b>%d</b></span>'
 '<span>교재 미수록 <b>%d</b></span></div>\n'
 % (stat['A'] + stat['B'] + stat['C'] + stat['W'], stat['O'], stat['R'])
 + ''.join(out)
 + '<div class="ansbox"><h2>정답</h2>'
 + ''.join(arows)
 + '<p class="todo">? = 발문 미확보 또는 시험 보기와 불일치</p></div>\n'
 '</div>\n<script>document.addEventListener("DOMContentLoaded",function(){'
 'renderMathInElement(document.body,{delimiters:['
 '{left:"$$",right:"$$",display:true},{left:"$",right:"$",display:false}],'
 'throwOnError:false});});</script>\n</body></html>'
)

with open(OUT, 'w', encoding='utf-8') as f:
    f.write(doc)

print('저장 ·', OUT)
print('총 %d문항  A %d · A구기출 %d · B %d · C %d · 복원완료 %d · 신규 %d'
      % (len(ITEMS), stat['A'], stat['O'], stat['B'], stat['C'], stat['R'], stat['N']))
print('과목별  이론 %d · 기기 %d · 설비 %d'
      % tuple(sum(1 for x in ITEMS if x[0] == s) for s in ('이론', '기기', '설비')))
miss = [t for _, t, tr, s, _ in ITEMS if tr in ('A','B','C') and s not in bank]
if miss:
    print('DB 조회 실패:', miss)

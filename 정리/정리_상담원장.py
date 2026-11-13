"""상담원장_2025.xlsx 정리 스크립트

결정 기록: d-001  (../my-wiki/decisions/d-001-상담원장-정리-결정.md)
  d-001-1  처리상태 «종결» → «완료», «대기» → «보류»            (CS기획팀 답)
  d-001-2  처리시간 음수는 값을 두고 맨 끝 _표시 칸에 이유를 적는다
  d-001-3  만족도 6~10점은 값을 두고 _표시 칸에 이유를 적는다 (평균 만족도에서는 뺀다)
  d-001-4  칸 이름은 원본 그대로 둔다 (맨 끝에 _표시 한 칸만 더한다)
  d-001-5  고객번호 앞뒤 공백을 없앤다

단계마다 근거를 적었다:
  [d-001-n]  결정 기록
  [작성기준]  원본 «작성기준» 시트에 적힌 것
  [대화]      2026-11-12 대화에서 사용자가 고른 것
  [짐작]      묻지 않고 정한 것. 아직 결정 기록 없음

원본은 읽기만 한다 (표준 라이브러리 zipfile 로 시트 XML을 읽는다).
쓰는 파일: <출력폴더>/상담원장_정리.csv · 지운줄.csv · 변경요약.md

사용: python 정리_상담원장.py [원본.xlsx] [출력폴더]
  기본값: ../../messy/상담원장_2025.xlsx · ../정리본   (이 파일 위치 기준)
"""
import csv
import datetime as dt
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent.parent / 'messy' / '상담원장_2025.xlsx'
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE.parent / '정리본'
SHEET = '원장'
FLAG = '_표시'                                     # [d-001-2·3·4] 맨 끝에 더하는 칸


# ── 원본 읽기 (읽기만) ─────────────────────────────────────────
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NSR = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NSP = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def col_index(ref):
    letters = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1, int(ref[len(letters):])


def read_sheet(path, name):
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rid = next(s.get(NSR + 'id') for s in wb.iter(NS + 'sheet') if s.get('name') == name)
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target = next(r.get('Target') for r in rels.iter(NSP + 'Relationship') if r.get('Id') == rid)
        target = target.lstrip('/')
        target = target if target.startswith('xl/') else 'xl/' + target
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
        cells = defaultdict(dict)
        with z.open(target) as f:
            for _, el in ET.iterparse(f):
                if el.tag != NS + 'c':
                    continue
                ci, ri = col_index(el.get('r'))
                t = el.get('t', 'n')
                ve = el.find(NS + 'v')
                if t == 'inlineStr':
                    val = ''.join(x.text or '' for x in el.iter(NS + 't'))
                elif ve is None:
                    val = None
                elif t == 's':
                    val = shared[int(ve.text)]
                elif t == 'str':
                    val = ve.text
                elif t == 'b':
                    val = ve.text == '1'
                else:
                    val = float(ve.text)
                    if val.is_integer():
                        val = int(val)
                cells[ri][ci] = val
                el.clear()
    head_row = min(cells)
    hdr = [cells[head_row][c] for c in sorted(cells[head_row])]
    rows = []
    for r in sorted(cells):
        if r == head_row:
            continue
        v = [cells[r].get(c) for c in range(len(hdr))]
        if any(x is not None for x in v):
            rows.append((r, v))
    return hdr, rows


# ── 바꾼 것 기록 ───────────────────────────────────────────────
LOG = []       # (원본행, 접수번호, 칸, 바꾼 것, 근거)


def note(r, key, col, what, basis):
    LOG.append((r, key, col, what, basis))


def pdate(s):
    for f in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'):
        try:
            return dt.datetime.strptime(s.strip(), f).date()
        except ValueError:
            pass
    raise ValueError(f'읽을 수 없는 날짜: {s!r}')


CH = {'전화': '전화', '유선': '전화', '전화상담': '전화', 'tel': '전화',
      '챗봇': '챗봇', 'chatbot': '챗봇', '앱': '앱', 'app': '앱', '모바일앱': '앱',
      '이메일': '이메일', 'e-mail': '이메일', '메일': '이메일'}
REG = {'경기도': '경기', '서울시': '서울', '서울특별시': '서울'}
STATUS = {'처리완료': ('완료', '[대화]+[짐작]'),        # 완료로 셀지는 [대화], 값을 바꿔 쓴 것은 [짐작]
          '종결': ('완료', '[d-001-1]'),
          '대기': ('보류', '[d-001-1]')}
RE = {'Y': 'Y', 'y': 'Y', '예': 'Y', 'TRUE': 'Y', 'N': 'N', 'n': 'N', '아니오': 'N', 'FALSE': 'N'}
PHONE_FROM = dt.date(2025, 7, 1)
WHY_NEG = '처리시간 음수 — 틀린 값이지만 진짜 값을 모름 (d-001-2)'
WHY_SAT = '만족도 1~5 밖(NPS와 같은 값) — NPS를 잘못 넣은 것으로 보이나 확인 못 함, 평균 만족도에서 뺌 (d-001-3)'


def main():
    hdr, data = read_sheet(SRC, SHEET)
    H = {h: i for i, h in enumerate(hdr)}
    assert FLAG not in hdr, f'원본에 이미 {FLAG} 칸이 있음'
    raw = {r: v for r, v in data}
    removed = {}                       # 원본행 → 지운 이유
    out = []
    for r, v in data:
        k = v[0]
        v = list(v)
        # 앞뒤 공백 [짐작] · 고객번호는 [d-001-5]
        for i, x in enumerate(v):
            if isinstance(x, str) and x != x.strip():
                note(r, k, hdr[i], '앞뒤 공백 제거', '[d-001-5]' if hdr[i] == '고객번호' else '[짐작]')
                v[i] = x.strip()
        # 테스트 접수 제외 [작성기준]
        if v[H['상담원ID']] == 'AG999':
            removed[r] = '테스트 접수 (상담원ID AG999)'
            continue
        # 접수일 YYYY-MM-DD [짐작]
        d = pdate(v[H['접수일']])
        if d.isoformat() != v[H['접수일']]:
            note(r, k, '접수일', '날짜 형식 YYYY-MM-DD', '[짐작]')
            v[H['접수일']] = d.isoformat()
        # 접수채널 넷으로 [짐작]
        c = v[H['접수채널']]
        n = CH[c.lower().replace(' ', '')]
        if n != c:
            note(r, k, '접수채널', '같은 채널 이름 통일', '[짐작]')
            v[H['접수채널']] = n
        # 상담원팀 가운데 공백 [짐작]
        t = v[H['상담원팀']]
        if ' ' in t:
            note(r, k, '상담원팀', '가운데 공백 제거', '[짐작]')
            v[H['상담원팀']] = t.replace(' ', '')
        # 지역 이름 [짐작]
        g = v[H['지역']]
        if g in REG:
            note(r, k, '지역', '같은 지역 이름 통일', '[짐작]')
            v[H['지역']] = REG[g]
        # 월정액 숫자 [짐작]
        m = v[H['월정액']]
        if isinstance(m, str):
            note(r, k, '월정액', '«원»·쉼표 떼고 숫자로', '[짐작]')
            v[H['월정액']] = int(m.replace(',', '').replace('원', ''))
        # 상담시간: 7/1 이후 전화는 초 → 분 [작성기준] (7/1 당일 포함 · 접수일로 판단은 [짐작])
        q = v[H['상담시간']]
        if v[H['접수채널']] == '전화' and d >= PHONE_FROM:
            note(r, k, '상담시간', '7/1 이후 전화: 초 → 분 (60으로 나눔)', '[작성기준]')
            q = q / 60
        v[H['상담시간']] = int(q) if float(q).is_integer() else round(q, 2)
        # 처리시간: 값은 그대로, 숫자 모양만 예전 정리본과 같게(소수 한 자리, 예: 4.0)
        if v[H['처리시간']] is not None:
            v[H['처리시간']] = float(v[H['처리시간']])
        # 처리상태
        p = v[H['처리상태']]
        if p in STATUS:
            new, basis = STATUS[p]
            note(r, k, '처리상태', f'{p} → {new}', basis)
            v[H['처리상태']] = new
        # 재문의 Y·N [짐작]
        y = v[H['재문의']]
        if y is not None and RE[y] != y:
            note(r, k, '재문의', 'Y/N 둘로 통일', '[짐작]')
            v[H['재문의']] = RE[y]
        # 만족도
        s = v[H['만족도']]
        if s is not None:
            if s in ('-', '미응답') or (not isinstance(s, str) and s == 0):
                note(r, k, '만족도', '미응답 표시(0·-·미응답) → 빈칸', '[작성기준]')
                s = None
            elif isinstance(s, str) and s.endswith('점'):
                note(r, k, '만족도', '«점» 떼고 숫자로', '[짐작]')
                s = int(s[:-1])
            # 6~10점은 값을 그대로 둔다 [d-001-3] — 아래 _표시
            v[H['만족도']] = s
        out.append((r, v))

    # 모든 칸이 같은 줄은 하나만 [작성기준: 12월 초 중복 추출] + [짐작: 판별 기준]
    first = {}
    kept = []
    for r, v in out:
        t = tuple(v)
        if t in first:
            removed[r] = f'중복 추출 ({first[t]}행과 모든 칸이 같음)'
            continue
        first[t] = r
        kept.append((r, v))
    # 같은 접수번호는 최종수정일이 가장 늦은 줄만 [작성기준]
    by_id = defaultdict(list)
    for r, v in kept:
        by_id[v[0]].append((r, v))
    latest = set()
    for k, lst in by_id.items():
        lst = sorted(lst, key=lambda x: x[1][H['최종수정일']])
        if len(lst) > 1 and lst[-1][1][H['최종수정일']] == lst[-2][1][H['최종수정일']]:
            raise SystemExit(f'최종수정일이 같은 줄이 있어 최신을 고를 수 없음: {k}')
        latest.add(lst[-1][0])
        for r, v in lst[:-1]:
            removed[r] = f'상태 변경 전 줄 ({lst[-1][0]}행이 최신, 최종수정일 {lst[-1][1][H["최종수정일"]]})'
    final = [(r, v) for r, v in kept if r in latest]

    # _표시 [d-001-2 · d-001-3]
    flags = {}
    for r, v in final:
        why = []
        pt = v[H['처리시간']]
        if pt is not None and pt < 0:
            why.append(WHY_NEG)
        s = v[H['만족도']]
        if s is not None and not (1 <= s <= 5):
            why.append(WHY_SAT)
        flags[r] = '; '.join(why)

    OUT.mkdir(parents=True, exist_ok=True)
    cell = lambda x: '' if x is None else x
    with open(OUT / '상담원장_정리.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(hdr + [FLAG])                            # [d-001-4] 칸 이름 그대로
        for r, v in final:
            w.writerow([cell(x) for x in v] + [flags[r]])
    with open(OUT / '지운줄.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['원본행'] + hdr + ['지운 이유'])
        for r, v in data:
            if r in removed:
                w.writerow([r] + [cell(x) for x in raw[r]] + [removed[r]])

    write_summary(hdr, H, data, final, removed, flags)


def write_summary(hdr, H, data, final, removed, flags):
    kept_rows = {r for r, _ in final}
    cnt_all, cnt_kept = Counter(), Counter()
    for r, k, col, what, basis in LOG:
        cnt_all[(basis, col, what)] += 1
        if r in kept_rows:
            cnt_kept[(basis, col, what)] += 1
    rm = Counter(why.split(' (')[0] for why in removed.values())
    order = ['테스트 접수', '중복 추출', '상태 변경 전 줄']
    rm_basis = {'테스트 접수': '[작성기준]', '중복 추출': '[작성기준]+[짐작]', '상태 변경 전 줄': '[작성기준]'}
    phone = [(r, v) for r, v in final if v[H['접수채널']] == '전화' and v[H['접수일']] >= PHONE_FROM.isoformat()]
    conv_rows = {r for r, k, col, what, b in LOG if col == '상담시간'}
    ph_vals = [v[H['상담시간']] for r, v in phone]
    other = [v[H['상담시간']] for r, v in final if r not in conv_rows]
    rawd = dict(data)
    raw_ph = sorted(rawd[r][H['상담시간']] for r, _ in phone)
    flagged = [(r, v) for r, v in final if flags[r]]

    L = []
    a = L.append
    a('# 변경요약 — 상담원장_2025.xlsx → 상담원장_정리.csv')
    a('')
    a('- 만든 스크립트: `정리/정리_상담원장.py`')
    a('- 결정 기록: d-001 (`../my-wiki/decisions/d-001-상담원장-정리-결정.md`)')
    a('- 근거 표시: [d-001-n] 결정 기록 · [작성기준] 원본 «작성기준» 시트 · [대화] 2026-11-12 대화에서 고른 것 · [짐작] 묻지 않고 정한 것')
    a('- 원본은 읽기만 했습니다.')
    a('')
    a('## 줄 수')
    a('')
    a('| 단계 | 근거 | 줄 수 |')
    a('|---|---|---|')
    a(f'| 원본 `{SHEET}` 시트 (머리글 빼고) | — | {len(data):,} |')
    for k in order:
        a(f'| − {k} | {rm_basis[k]} | −{rm[k]:,} |')
    a(f'| = 정리본 | — | **{len(final):,}** |')
    a('')
    a(f'빼기 확인: {len(data):,} − {sum(rm.values()):,} = {len(data) - sum(rm.values()):,} (정리본 {len(final):,}). 지운 줄은 모두 `지운줄.csv`에 있고, 맨 끝 칸에 지운 이유가 있습니다.')
    a('')
    a('## 칸 값을 바꾼 것')
    a('')
    a('«처리한 줄»은 줄을 빼기 전에 바꾼 칸 수, «정리본에 남은 줄»은 그중 정리본에 남은 줄의 칸 수입니다.')
    a('')
    a('| 근거 | 칸 | 바꾼 것 | 처리한 줄 | 정리본에 남은 줄 |')
    a('|---|---|---|---|---|')
    for (basis, col, what), n in sorted(cnt_all.items(), key=lambda x: (hdr.index(x[0][1]), -x[1])):
        a(f'| {basis} | {col} | {what} | {n:,} | {cnt_kept[(basis, col, what)]:,} |')
    a('')
    a('## 상담시간 — 7월 이후 전화 줄은 초에서 분으로 바뀌었습니다')
    a('')
    a('- 근거: [작성기준] «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위로 기록됨».')
    a('- 고른 방식 [짐작]: 접수일로 판단했고 7월 1일 당일도 넣었습니다. 접수채널은 넷으로 묶은 뒤의 «전화» (유선 · 전화상담 · TEL · 앞 공백 « 전화» 포함).')
    a(f'- 바꾼 줄: 처리한 줄 {cnt_all[("[작성기준]", "상담시간", "7/1 이후 전화: 초 → 분 (60으로 나눔)")]:,} · 정리본에 남은 줄 {len(phone):,}.')
    a(f'- 정리본에 남은 {len(phone):,}줄: 바꾸기 전(초) {raw_ph[0]:g} ~ {raw_ph[-1]:g}, 중앙값 {median(raw_ph):g} → 바꾼 뒤(분) {min(ph_vals):g} ~ {max(ph_vals):g}, 중앙값 {median(ph_vals):g}.')
    a(f'- 그중 60으로 나눠 `.5`가 된 줄 {sum(1 for x in ph_vals if not float(x).is_integer()):,} (반올림하지 않았습니다 [짐작]).')
    a(f'- 바꾸지 않은 줄(분 단위 그대로) {len(other):,}: {min(other):g} ~ {max(other):g}분, 중앙값 {median(other):g}.')
    a('- **이 칸의 값은 이제 모두 분입니다.** 정리본 줄의 원래 초 값은 원본 파일에서만 볼 수 있습니다.')
    a('')
    a(f'## `{FLAG}` 칸 [d-001-2 · d-001-3]')
    a('')
    a(f'- 맨 끝에 `{FLAG}` 칸을 더했습니다. 기존 {len(hdr)}칸의 이름과 순서는 원본 그대로입니다 [d-001-4].')
    a(f'- 표시한 줄 {len(flagged)}: 처리시간 음수 {sum(1 for r, _ in flagged if WHY_NEG in flags[r])} · 만족도 6~10점 {sum(1 for r, _ in flagged if WHY_SAT in flags[r])} · 둘 다 {sum(1 for r, _ in flagged if WHY_NEG in flags[r] and WHY_SAT in flags[r])}. 값은 바꾸지 않았습니다.')
    a('')
    a('| 접수번호 | 원본 행 | 처리시간 | 만족도 | NPS | 이유 |')
    a('|---|---|---|---|---|---|')
    for r, v in flagged:
        a(f'| {v[0]} | {r} | {v[H["처리시간"]]} | {cell_md(v[H["만족도"]])} | {cell_md(v[H["NPS"]])} | {flags[r]} |')
    a('')
    a('## 바꾸지 않은 것')
    a('')
    a('- 결정 기록에 없는 처리는 더하지 않았습니다. 처리중 · 진행중은 따로 둡니다. 미완료인데 만족도나 처리시간이 있는 줄도 그대로입니다.')
    (OUT / '변경요약.md').write_text('\n'.join(L) + '\n', encoding='utf-8')


def median(xs):
    s = sorted(xs)
    n = len(s)
    return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2


def cell_md(x):
    return '' if x is None else x


if __name__ == '__main__':
    main()

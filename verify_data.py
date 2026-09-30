"""
업데이트 전에 돌리는 데이터 이상 점검 스크립트. 문제가 있으면 [경고]로 출력(자동 수정은 안 함).

점검 항목:
  1. 어제 있던 단지가 오늘 빠짐 (집셀 파일 누락)
  2. 단지별 매매 매물 수 급변 (30% 이상 & 3건 이상)
  3. 어제는 중복매물이 있었는데 오늘은 전부 단독 → 집셀 다운로드 설정 문제 의심
  4. 어제 전월세가 3건 이상 있었는데 오늘 0건 → "매매만"으로 받아졌을 가능성
  5. 같은 매물번호가 서로 다른 매물에 동시에 등장
  6. 소멸+신규에 동·타입·가격이 같은 매물이 동시에 있음(층 표기만 바뀐 재등록 의심)

사용법: python3 verify_data.py <어제_index.html> <오늘_all_deals.json> <snapshot_diff.json> <rent_diff.json> <어제날짜> <갱신된_archive.json>
"""
import json, sys
from collections import Counter

def extract_const(html, name):
    s = open(html, encoding='utf-8').read()
    i = s.find(f'const {name} = ') + len(f'const {name} = ')
    return json.JSONDecoder().raw_decode(s[i:])[0]

def main():
    html, deals_path, sd_path, rd_path, prev_date, archive_path = sys.argv[1:7]
    deals = json.load(open(deals_path, encoding='utf-8'))
    archive = json.load(open(archive_path, encoding='utf-8'))
    prev = archive[prev_date]
    old_all = extract_const(html, 'LISTINGS_ALL')
    warn = 0
    def W(msg):
        nonlocal warn; warn += 1; print('[경고]', msg)

    pc = Counter((e['type'], e['complex']) for e in prev)
    nc = Counter((e['type'], e['complex']) for e in deals['sale'])
    for k in pc:
        if k not in nc: W(f'오늘 파일 없음: {k} (어제 매매 {pc[k]}건)')
    for k in nc:
        o, n = pc.get(k, 0), nc[k]
        if o and abs(n - o) >= 3 and abs(n - o) / o >= 0.3: W(f'매물 수 급변: {k} {o}→{n}')

    for k in nc:
        y = [e for e in prev if (e['type'], e['complex']) == k]
        t = [e for e in deals['sale'] if (e['type'], e['complex']) == k]
        if len(t) >= 5 and sum(len(e['idSet']) > 1 for e in y) >= 3 and all(len(e['listingIds']) == 1 for e in t):
            W(f'중복매물 0건: {k} (어제 중복그룹 {sum(len(e["idSet"]) > 1 for e in y)}개) — 다운로드 설정 확인')

    for dt, key in [('전세', 'jeonse'), ('월세', 'wolse')]:
        o = Counter((e['type'], e['complex']) for e in old_all if e.get('dealType') == dt)
        n = Counter((e['type'], e['complex']) for e in deals[key])
        for k in o:
            if o[k] >= 3 and not n.get(k) and k in nc: W(f'{dt} 0건: {k} (어제 {o[k]}건)')

    ids = Counter(i for e in deals['sale'] for i in e['listingIds'])
    dup = [i for i, c in ids.items() if c > 1]
    if dup: W(f'매물번호 중복 등장 {len(dup)}개: {dup[:5]}')

    for label, path, gk, fk in [('매매', sd_path, 'gone', 'fresh'), ('전월세', rd_path, 'gone', 'new')]:
        d = json.load(open(path, encoding='utf-8'))
        sig = lambda e: (e['type'], e['complex'], e.get('dealType'), e['building'], e['unitType'], e['price'], e.get('rent'))
        g = Counter(sig(e) for e in d[gk]); f = Counter(sig(e) for e in d[fk])
        for k in g:
            if k in f: W(f'{label} 소멸·신규 동시(층 표기 변경 재등록 의심): {k}')

    print(f'점검 완료 — 경고 {warn}건')

if __name__ == '__main__':
    main()

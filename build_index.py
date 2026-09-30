"""
어제 index.html에 오늘 만든 데이터들을 넣어서 새 index.html을 만든다.
(화면 코드는 그대로 두고 `const 이름 = {...}` 데이터 부분만 교체)

사용법:
  python3 build_index.py <어제_index.html> <새_index.html> <오늘날짜 YYYY-MM-DD> \
    --deals all_deals.json --snapshot-diff X_snapshot_diff.json --persist-gone X_persist_gone.json \
    --real-tx real_tx.json --location real_tx_location.json \
    --sale-real sale_real.json --complex-stats sale_complexstats.json --rent-stats rent.json \
    --new-tx new_tx.json --new-rent-tx new_rent_tx.json --rent-diff rent_diff.json

교체 대상:
  LISTINGS_ALL(매매+전세+월세), APT_LISTINGS(구/동별 매물 요약), SNAPSHOT_DIFF, PERSIST_GONE,
  REAL_TX, COMPLEX_LOCATION, REAL, COMPLEX_STATS, REAL_RENT, NEW_TX, NEW_RENT_TX, RENT_DIFF,
  ALL_UNIT_TYPES/ALL_BUILDINGS/ALL_DIRECTIONS(기존 값 + 오늘 매매 매물 합집합 — 오늘 매물이 없어도
  예전 타입이 드롭다운에서 안 사라지게), DATA_LAST_UPDATED.
건드리지 않는 것: EXISTING_MAX(1년치 국토부 파일로는 재현 불가한 과거 최고가 기준점), COLORS, 화면 코드.
"""
import json, argparse
from collections import defaultdict

def find_const(s, name):
    start = s.find(f'const {name} = ')
    if start < 0:
        raise KeyError(name)
    vstart = start + len(f'const {name} = ')
    _, length = json.JSONDecoder().raw_decode(s[vstart:])
    return vstart, vstart + length

def replace_const(s, name, obj):
    a, b = find_const(s, name)
    return s[:a] + json.dumps(obj, ensure_ascii=False) + s[b:]

def get_const(s, name):
    a, b = find_const(s, name)
    return json.loads(s[a:b])

def build_apt_listings(sale):
    out = {}
    groups = defaultdict(list)
    for e in sale:
        groups[(e['gu'], e['type'], e['dong'])].append(e)
    for (gu, t, dong), items in sorted(groups.items()):
        gap = sum(1 for e in items if e.get('gap'))
        recent = sorted(items, key=lambda e: e.get('checkedDate') or '', reverse=True)[:8]
        out.setdefault(gu, {}).setdefault(t, {})[dong] = {
            'count': len(items),
            'avgPrice': round(sum(e['price'] for e in items) / len(items), 2),
            'gapCount': gap,
            'gapPct': round(gap / len(items) * 100, 1),
            'recent': [{'complex': e['complex'], 'floor': e['floor'], 'type': e['unitType'],
                        'area': e['area'], 'price': e['price'], 'direction': e['direction'],
                        'checkedDate': e.get('checkedDate'), 'gap': e.get('gap', False)} for e in recent],
        }
    return out

def merge_options(existing, sale, field):
    out = {t: {cx: {k: set(v) for k, v in areas.items()} for cx, areas in m.items()} for t, m in existing.items()}
    for e in sale:
        val = e.get(field)
        if val in (None, '', 'None'):
            continue
        out.setdefault(e['type'], {}).setdefault(e['complex'], {}).setdefault(str(int(e['area'])), set()).add(str(val))
    return {t: {cx: {k: sorted(v) for k, v in areas.items()} for cx, areas in m.items()} for t, m in out.items()}

def main():
    p = argparse.ArgumentParser()
    p.add_argument('old_html'); p.add_argument('new_html'); p.add_argument('date')
    for a in ['deals', 'snapshot-diff', 'persist-gone', 'real-tx', 'location', 'sale-real',
              'complex-stats', 'rent-stats', 'new-tx', 'new-rent-tx', 'rent-diff']:
        p.add_argument('--' + a, required=True)
    args = p.parse_args()
    load = lambda f: json.load(open(f, encoding='utf-8'))

    s = open(args.old_html, encoding='utf-8').read()
    deals = load(args.deals)
    sale = deals['sale']

    s = replace_const(s, 'LISTINGS_ALL', sale + deals['jeonse'] + deals['wolse'])
    s = replace_const(s, 'APT_LISTINGS', build_apt_listings(sale))
    s = replace_const(s, 'SNAPSHOT_DIFF', load(args.snapshot_diff))
    s = replace_const(s, 'PERSIST_GONE', load(args.persist_gone))
    s = replace_const(s, 'REAL_TX', load(args.real_tx))
    s = replace_const(s, 'COMPLEX_LOCATION', load(args.location))
    s = replace_const(s, 'REAL', load(args.sale_real))
    s = replace_const(s, 'COMPLEX_STATS', load(args.complex_stats))
    s = replace_const(s, 'REAL_RENT', load(args.rent_stats))
    s = replace_const(s, 'NEW_TX', load(args.new_tx))
    s = replace_const(s, 'NEW_RENT_TX', load(args.new_rent_tx))
    s = replace_const(s, 'RENT_DIFF', load(args.rent_diff))
    for name, field in [('ALL_UNIT_TYPES', 'unitType'), ('ALL_BUILDINGS', 'building'), ('ALL_DIRECTIONS', 'direction')]:
        s = replace_const(s, name, merge_options(get_const(s, name), sale, field))

    marker = "const DATA_LAST_UPDATED = '"
    i = s.find(marker) + len(marker)
    j = s.find("'", i)
    s = s[:i] + args.date + s[j:]

    open(args.new_html, 'w', encoding='utf-8').write(s)
    print(f"{args.new_html} 생성 ({len(s):,} bytes), DATA_LAST_UPDATED={args.date}")

if __name__ == '__main__':
    main()

"""
매물 변동추이(SNAPSHOT_DIFF)·지속소멸매물(PERSIST_GONE) 재계산 스크립트.

핵심 아이디어: 대화가 바뀌어도 데이터가 안 끊기게, 상세정보(중개사·확인일자·매물번호 등)까지
전부 담은 "완전 보관함"(listing_full_archive.json)을 깃허브에 계속 쌓아두고, 이 스크립트가
그걸 읽어서 새 날짜를 추가한 뒤 SNAPSHOT_DIFF·PERSIST_GONE을 다시 계산한다.

사용법:
  python3 update_archive_and_recompute.py \
    --archive listing_full_archive.json \
    --new-date 2026-09-09 \
    --new-deals all_deals_0909.json \
    --out-prefix out_0909

--new-deals는 extract_all_deals.py의 출력 파일(그 안의 'sale' 배열만 사용).
--archive에 --new-date가 이미 있으면 그 날짜만 덮어씀(당일 재업로드 대응).

출력:
  <out-prefix>_archive.json      : 갱신된 전체 보관함 (다음날 이걸 다시 --archive로 넘기면 됨)
  <out-prefix>_snapshot_diff.json: index.html의 SNAPSHOT_DIFF에 그대로 붙여넣을 JSON
  <out-prefix>_persist_gone.json : index.html의 PERSIST_GONE에 그대로 붙여넣을 JSON
"""
import json, argparse
from datetime import datetime

def key(d):
    return (d['complex'], d['building'], d['floor'], d['unitType'])

def to_archive_entry(d):
    return {'gu': d['gu'], 'type': d['type'], 'dong': d['dong'], 'complex': d['complex'],
            'building': d['building'], 'floor': d['floor'], 'unitType': d['unitType'],
            'area': d['area'], 'price': d['price'], 'direction': d['direction'],
            'checkedDate': d.get('checkedDate'), 'gap': d.get('gap', False),
            'agencies': d.get('agencies', []), 'idSet': d.get('listingIds', d.get('idSet', []))}

def main():
    p = argparse.ArgumentParser()
    p.add_argument('--archive', required=True)
    p.add_argument('--new-date', required=True, help='YYYY-MM-DD')
    p.add_argument('--new-deals', required=True)
    p.add_argument('--out-prefix', required=True)
    args = p.parse_args()

    archive = json.load(open(args.archive, encoding='utf-8'))
    new_deals = json.load(open(args.new_deals, encoding='utf-8'))
    new_sale = new_deals['sale'] if 'sale' in new_deals else new_deals

    # 오늘 터치된 단지만 새 데이터로 교체, 나머지는 어제(가장 최근 날짜) 걸 그대로 이어받음
    dates_sorted = sorted(archive.keys())
    prev_date = dates_sorted[-1] if dates_sorted else None
    touched_today = {(d['type'], d['complex']) for d in new_sale}

    if prev_date:
        carried = [e for e in archive[prev_date] if (e['type'], e['complex']) not in touched_today]
    else:
        carried = []
    today_entries = [to_archive_entry(d) for d in new_sale]
    archive[args.new_date] = carried + today_entries

    # ── SNAPSHOT_DIFF: 최근 두 날짜(직전 vs 오늘) 비교 ──
    dates_sorted = sorted(archive.keys())
    snapshot_diff = {'gone': [], 'fresh': [], 'changed': [], 'oldDate': None, 'newDate': args.new_date}
    if len(dates_sorted) >= 2:
        old_date = dates_sorted[-2]
        old_map = {key(e): e for e in archive[old_date]}
        new_map = {key(e): e for e in archive[args.new_date]}
        fresh_keys = set(new_map) - set(old_map)
        gone_keys = set(old_map) - set(new_map)
        common_keys = set(old_map) & set(new_map)
        snapshot_diff['gone'] = [old_map[k] for k in gone_keys]
        snapshot_diff['fresh'] = [new_map[k] for k in fresh_keys]
        for k in common_keys:
            o, n = old_map[k], new_map[k]
            if o['price'] != n['price']:
                snapshot_diff['changed'].append({
                    'complex': o['complex'], 'type': o['type'], 'building': o['building'],
                    'floor': o['floor'], 'unitType': o['unitType'],
                    'oldPrice': o['price'], 'newPrice': n['price'], 'diff': round(n['price']-o['price'], 2),
                })
        snapshot_diff['oldDate'] = old_date

    # ── PERSIST_GONE: 마지막으로 보인 날짜가 "직전 날짜"보다도 더 예전인 매물 (최소 한 번의 비교 주기 이상 계속 없음) ──
    last_seen = {}
    for d in dates_sorted:
        for e in archive[d]:
            last_seen[key(e)] = (d, e)
    persist_gone_list = []
    if len(dates_sorted) >= 2:
        cutoff_date = dates_sorted[-2]  # 직전 날짜에도 없었으면 "지속 소멸"
        latest_keys = {key(e) for e in archive[args.new_date]}
        for k, (d, e) in last_seen.items():
            if k in latest_keys:
                continue
            if d < cutoff_date:
                persist_gone_list.append({
                    'complex': e['complex'], 'type': e['type'], 'dong': e['dong'],
                    'building': e['building'], 'floor': e['floor'], 'unitType': e['unitType'],
                    'area': e['area'], 'price': e['price'], 'direction': e['direction'],
                    'lastSeenDate': d.replace('-', '.'),
                })
    persist_gone = {
        'range': {'from': dates_sorted[0].replace('-', '.') if dates_sorted else '', 'to': args.new_date.replace('-', '.')},
        'list': persist_gone_list,
    }

    json.dump(archive, open(f'{args.out_prefix}_archive.json', 'w', encoding='utf-8'), ensure_ascii=False)
    json.dump(snapshot_diff, open(f'{args.out_prefix}_snapshot_diff.json', 'w', encoding='utf-8'), ensure_ascii=False)
    json.dump(persist_gone, open(f'{args.out_prefix}_persist_gone.json', 'w', encoding='utf-8'), ensure_ascii=False)

    print(f"보관함 날짜 수: {len(archive)} ({dates_sorted[0]} ~ {args.new_date})")
    print(f"SNAPSHOT_DIFF: 신규 {len(snapshot_diff['fresh'])} / 소멸 {len(snapshot_diff['gone'])} / 가격변동 {len(snapshot_diff['changed'])}")
    print(f"PERSIST_GONE: {len(persist_gone_list)}건, 범위 {persist_gone['range']}")

if __name__ == '__main__':
    main()

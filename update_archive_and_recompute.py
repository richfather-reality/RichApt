"""
매물 변동추이(SNAPSHOT_DIFF)·지속소멸매물(PERSIST_GONE) 재계산 스크립트.

핵심 아이디어: 대화가 바뀌어도 데이터가 안 끊기게, 상세정보(중개사·확인일자·매물번호 등)까지
전부 담은 "완전 보관함"(listing_full_archive.json)을 깃허브에 계속 쌓아두고, 이 스크립트가
그걸 읽어서 새 날짜를 추가한 뒤 SNAPSHOT_DIFF·PERSIST_GONE을 다시 계산한다.

매칭 로직(중요): (단지,동,층,타입)만으로는 같은 표기 안에 실제로는 서로 다른 여러 집이
섞여있는 경우가 흔함(예: 같은 층에 같은 타입이 여러 호수 있음). 그래서 키 하나에 항목 하나가
아니라 리스트로 모아두고, 그 안에서 아래 순서로 짝을 맞춘다:
  1차: 매물번호(idSet) 겹치는 것끼리 (대표가 바뀌어도 중복 목록에 예전 매물번호가 남아있으면 잡힘)
  2차: 그래도 안 맞으면 가격이 정확히 같은 것끼리 (매물번호가 통째로 바뀐 경우)
  3차: 그래도 남았는데 옛날 개수와 오늘 개수가 똑같으면, 시장에서 뭔가 빠지거나 늘어난 게
       아니라 그냥 재등록된 걸로 보고 가격순으로 짝지어 가격변동으로 처리
       (개수가 그대로인데 소멸+신규로 쪼개면 실제로 안 바뀐 매물이 나갔다 들어온 것처럼
       과장되게 보임 — 기흥역세권 자료와 비교해서 확인된 방식)
개수가 안 맞을 때만 진짜로 소멸/신규로 본다.

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
from collections import defaultdict
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

    dates_sorted = sorted(archive.keys())
    prev_date = dates_sorted[-1] if dates_sorted else None
    touched_today = {(d['type'], d['complex']) for d in new_sale}

    if prev_date:
        carried = [e for e in archive[prev_date] if (e['type'], e['complex']) not in touched_today]
    else:
        carried = []
    today_entries = [to_archive_entry(d) for d in new_sale]
    archive[args.new_date] = carried + today_entries

    dates_sorted = sorted(archive.keys())
    snapshot_diff = {'gone': [], 'fresh': [], 'changed': [], 'oldDate': None, 'newDate': args.new_date}
    if len(dates_sorted) >= 2:
        old_date = dates_sorted[-2]
        old_by_key = defaultdict(list)
        for e in archive[old_date]: old_by_key[key(e)].append(e)
        new_by_key = defaultdict(list)
        for e in archive[args.new_date]: new_by_key[key(e)].append(e)

        unmatched_old, unmatched_new = [], []
        for k in set(old_by_key) | set(new_by_key):
            olds, news = list(old_by_key.get(k, [])), list(new_by_key.get(k, []))
            if len(olds) == 1 and len(news) == 1:
                o, n = olds[0], news[0]
                if o['price'] != n['price']:
                    snapshot_diff['changed'].append({
                        'complex': o['complex'], 'type': o['type'], 'building': n['building'],
                        'floor': n['floor'], 'unitType': n['unitType'],
                        'oldPrice': o['price'], 'newPrice': n['price'], 'diff': round(n['price']-o['price'], 2),
                    })
                continue
            used = [False]*len(news)
            for o in olds:
                o_ids = set(o.get('idSet', []))
                found = None
                for j, n in enumerate(news):
                    if used[j]: continue
                    if o_ids & set(n.get('idSet', [])):
                        found = j; break
                if found is not None:
                    used[found] = True
                    n = news[found]
                    if o['price'] != n['price']:
                        snapshot_diff['changed'].append({
                            'complex': o['complex'], 'type': o['type'], 'building': n['building'],
                            'floor': n['floor'], 'unitType': n['unitType'],
                            'oldPrice': o['price'], 'newPrice': n['price'], 'diff': round(n['price']-o['price'], 2),
                        })
                    o['_matched'] = True
            for o in olds:
                if o.get('_matched'): continue
                found = None
                for j, n in enumerate(news):
                    if used[j]: continue
                    if o['price'] == n['price']:
                        found = j; break
                if found is not None:
                    used[found] = True
                    o['_matched'] = True
            remaining_old = [o for o in olds if not o.get('_matched')]
            remaining_new = [n for j, n in enumerate(news) if not used[j]]
            if len(remaining_old) == len(remaining_new) and remaining_old:
                for o, n in zip(sorted(remaining_old, key=lambda x: x['price']),
                                 sorted(remaining_new, key=lambda x: x['price'])):
                    if o['price'] != n['price']:
                        snapshot_diff['changed'].append({
                            'complex': o['complex'], 'type': o['type'], 'building': n['building'],
                            'floor': n['floor'], 'unitType': n['unitType'],
                            'oldPrice': o['price'], 'newPrice': n['price'], 'diff': round(n['price']-o['price'], 2),
                        })
                    o['_matched'] = True
                for j in range(len(news)):
                    used[j] = True
            else:
                for o in remaining_old:
                    unmatched_old.append(o)
            for j, n in enumerate(news):
                if not used[j]: unmatched_new.append(n)

        gone_by_complex = defaultdict(list)
        for o in unmatched_old: gone_by_complex[o['complex']].append(o)
        fresh_by_complex = defaultdict(list)
        for n in unmatched_new: fresh_by_complex[n['complex']].append(n)
        consumed_fresh = set()
        final_gone = []
        for cx, olds in gone_by_complex.items():
            candidates = fresh_by_complex.get(cx, [])
            for o in olds:
                o_ids = set(o.get('idSet', []))
                match = None
                if o_ids:
                    for idx, n in enumerate(candidates):
                        if idx in consumed_fresh: continue
                        if o_ids & set(n.get('idSet', [])) and o['price'] == n['price']:
                            match = idx; break
                if match is not None:
                    consumed_fresh.add(match)
                else:
                    final_gone.append(o)
        final_fresh = []
        for cx, news in fresh_by_complex.items():
            olds = gone_by_complex.get(cx, [])
            local_consumed = set()
            for o in olds:
                o_ids = set(o.get('idSet', []))
                if not o_ids: continue
                for idx, n in enumerate(news):
                    if idx in local_consumed: continue
                    if o_ids & set(n.get('idSet', [])) and o['price'] == n['price']:
                        local_consumed.add(idx); break
            for idx, n in enumerate(news):
                if idx not in local_consumed:
                    final_fresh.append(n)

        snapshot_diff['gone'] = final_gone
        snapshot_diff['fresh'] = final_fresh
        snapshot_diff['oldDate'] = old_date

    latest_id_prices = defaultdict(set)
    for e in archive[args.new_date]:
        for i in e.get('idSet', []):
            latest_id_prices[i].add(e['price'])

    COMPLEX_TRACKING_RESET = {
        '강남마을6단지자연앤': '2026-09-17',
    }

    tracked_units = []
    prev_date_walk = None
    for d in dates_sorted:
        by_key_today = defaultdict(list)
        for e in archive[d]:
            reset_date = COMPLEX_TRACKING_RESET.get(e['complex'])
            if reset_date and d < reset_date:
                continue
            by_key_today[key(e)].append(e)

        used_today = defaultdict(set)
        for unit in tracked_units:
            candidates = by_key_today.get(unit['key'], [])
            found = None
            for idx, cand in enumerate(candidates):
                if idx in used_today[unit['key']]: continue
                if unit['ids'] & set(cand.get('idSet', [])) or unit['price'] == cand['price']:
                    found = idx; break
            if found is not None:
                used_today[unit['key']].add(found)
                cand = candidates[found]
                unit['last_date'] = d
                unit['entry'] = cand
                unit['ids'] = unit['ids'] | set(cand.get('idSet', []))
                unit['price'] = cand['price']

        if prev_date_walk is not None:
            pending_by_key = defaultdict(list)
            for unit in tracked_units:
                if unit['last_date'] != prev_date_walk: continue
                k = unit['key']
                cands = by_key_today.get(k, [])
                if len(used_today[k]) >= len(cands): continue
                pending_by_key[k].append(unit)
            for k, pending_units in pending_by_key.items():
                cands = by_key_today.get(k, [])
                remaining_cands = [idx for idx in range(len(cands)) if idx not in used_today[k]]
                if len(pending_units) == len(remaining_cands) and pending_units:
                    for unit, idx in zip(sorted(pending_units, key=lambda u: u['price']),
                                          sorted(remaining_cands, key=lambda i: cands[i]['price'])):
                        used_today[k].add(idx)
                        cand = cands[idx]
                        unit['last_date'] = d
                        unit['entry'] = cand
                        unit['ids'] = unit['ids'] | set(cand.get('idSet', []))
                        unit['price'] = cand['price']

        for k, candidates in by_key_today.items():
            for idx, cand in enumerate(candidates):
                if idx not in used_today[k]:
                    tracked_units.append({'key': k, 'last_date': d, 'entry': cand,
                                           'ids': set(cand.get('idSet', [])), 'price': cand['price']})
        prev_date_walk = d

    persist_gone_list = []
    if len(dates_sorted) >= 2:
        cutoff_date = dates_sorted[-2]
        for unit in tracked_units:
            if unit['last_date'] == args.new_date:
                continue
            e = unit['entry']
            still_alive = any(e['price'] in latest_id_prices.get(i, set()) for i in e.get('idSet', []))
            if still_alive:
                continue
            if unit['last_date'] < cutoff_date:
                persist_gone_list.append({
                    'complex': e['complex'], 'type': e['type'], 'dong': e['dong'],
                    'building': e['building'], 'floor': e['floor'], 'unitType': e['unitType'],
                    'area': e['area'], 'price': e['price'], 'direction': e['direction'],
                    'lastSeenDate': unit['last_date'].replace('-', '.'),
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

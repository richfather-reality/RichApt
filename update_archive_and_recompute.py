"""
매물 변동추이(SNAPSHOT_DIFF)·지속소멸매물(PERSIST_GONE) 재계산 스크립트.
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
        # 같은 (단지,동,층,타입) 키 안에 실제로는 서로 다른 집이 여러 개 있는 경우가 있음
        # (예: "고/44"라는 층 표기 하나에 진짜 다른 두 집이 같이 묶임). 그래서 키 하나에 항목 하나가
        # 아니라 리스트로 모아두고, 그 안에서 매물번호(idSet)로 짝을 맞춰 어떤 게 남고 어떤 게
        # 없어졌는지 가려냄.
        old_by_key = defaultdict(list)
        for e in archive[old_date]: old_by_key[key(e)].append(e)
        new_by_key = defaultdict(list)
        for e in archive[args.new_date]: new_by_key[key(e)].append(e)

        unmatched_old, unmatched_new = [], []
        for k in set(old_by_key) | set(new_by_key):
            olds, news = list(old_by_key.get(k, [])), list(new_by_key.get(k, []))
            # 흔한 경우(키 하나에 옛날도 하나, 오늘도 하나)는 그냥 같은 매물로 봄 — 매물번호는
            # 다른 중개사가 새로 올리면 자연스럽게 바뀔 수 있어서, 이 경우까지 idSet으로
            # 재확인하면 오히려 정상적인 "가격만 바뀐" 케이스를 소멸+신규로 잘못 쪼개버림.
            # 키 하나에 여러 개가 겹칠 때만(위 "고/44" 사례처럼) idSet으로 구분함.
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
            # 1차: 매물번호(idSet) 겹치는 것끼리 짝짓기
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
            # 2차: 매물번호가 하나도 안 겹쳐도(중개사가 통째로 바뀐 경우), 가격이 정확히 같으면
            # 같은 집으로 봄 — 같은 키(동일 층표기·타입) 안에 가격까지 같은 게 우연히 여럿일 확률은
            # 낮아서 안전한 매칭임.
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
                else:
                    unmatched_old.append(o)
            for j, n in enumerate(news):
                if not used[j]: unmatched_new.append(n)

        # 위에서 짝을 못 찾은 것들 중에는 "표기(층 등)가 통째로 바뀌어서 키 자체가 달라진" 같은 매물이
        # 섞여있을 수 있음(예: "고/47"->"32/47"). 매물번호가 겹치고 가격도 같을 때만 같은 매물로 봄
        # — 가격이 다르면 표기만 바뀐 게 아니라 그냥 다른 매물이 매물번호를 우연히 공유하는 경우일
        # 수 있어서(예: 한 중개사가 같은 단지 여러 집을 동시에 광고) 소멸/신규 그대로 둠.
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
        # 단지별로 소멸-신규 매칭에서 못 짝지어진 신규만 최종 "신규"로 남김
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

    # "매물번호 -> 그 매물번호를 포함한 최신 매물들의 가격 목록"으로 만들어서,
    # 가격까지 같을 때만 "여전히 살아있는 매물"로 봄 (매물번호만 겹치고 가격이 다르면 다른 매물).
    latest_id_prices = defaultdict(set)
    for e in archive[args.new_date]:
        for i in e.get('idSet', []):
            latest_id_prices[i].add(e['price'])

    # 지속소멸도 SNAPSHOT_DIFF와 같은 문제(같은 키 안에 실제로는 여러 채가 섞여있음)를 그대로 갖고
    # 있었음 — 그것도 이틀이 아니라 전체 날짜에 걸쳐 누적되는 거라 영향이 더 클 수 있음. 그래서 키
    # 하나당 항목 하나만 기억하는 대신, 날짜 순으로 훑으면서 "그 키 안의 여러 채" 각각을 매물번호
    # (안 겹치면 가격)로 이어붙여 계속 추적함.
    COMPLEX_TRACKING_RESET = {
        '강남마을6단지자연앤': '2026-09-17',
    }

    tracked_units = []  # 각 원소: {'key','last_date','entry','ids','price'}
    for d in dates_sorted:
        by_key_today = defaultdict(list)
        for e in archive[d]:
            reset_date = COMPLEX_TRACKING_RESET.get(e['complex'])
            if reset_date and d < reset_date:
                continue
            by_key_today[key(e)].append(e)

        used_today = defaultdict(set)  # key -> {인덱스,...}
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

        for k, candidates in by_key_today.items():
            for idx, cand in enumerate(candidates):
                if idx not in used_today[k]:
                    tracked_units.append({'key': k, 'last_date': d, 'entry': cand,
                                           'ids': set(cand.get('idSet', [])), 'price': cand['price']})

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

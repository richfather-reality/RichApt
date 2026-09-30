"""
전세/월세 매물변동(RENT_DIFF: 신규/소멸/가격변동) 생성 스크립트.
어제 index.html의 LISTINGS_ALL 안 전세·월세 매물과, 오늘 extract_all_deals.py 결과(jeonse/wolse)를 비교한다.
(전월세는 listing_full_archive.json에 안 쌓이므로, 어제 기준은 항상 "직전 index.html"에서 가져옴)

매칭 방식은 매매(update_archive_and_recompute.py)와 같은 원칙:
  같은 (주택유형, 단지, 전세/월세, 동, 층, 타입) 묶음 안에서
  1차: 매물번호(listingIds) 겹치는 것끼리
  2차: 보증금+월세가 정확히 같은 것끼리
  3차: 남은 개수가 어제/오늘 같으면 재등록으로 보고 보증금 순으로 짝지어 가격변동 처리
  개수가 다를 때만 진짜 소멸/신규.
가격변동(changed)은 어제 매물 정보를 기본으로 oldPrice/newPrice/oldRent/newRent를 붙임(기존 화면 형식).

사용법:
  python3 build_rent_diff.py <어제_index.html> <오늘_all_deals.json> <어제날짜 YYYY-MM-DD> <오늘날짜> <출력.json>
"""
import json, sys
from collections import defaultdict

def extract_const(html, name):
    s = open(html, encoding='utf-8').read()
    i = s.find(f'const {name} = ') + len(f'const {name} = ')
    return json.JSONDecoder().raw_decode(s[i:])[0]

def gkey(e):
    return (e['type'], e['complex'], e['dealType'], e['building'], e['floor'], e['unitType'])

def money(e):
    return (e['price'], e.get('rent') or 0)

def main():
    html, deals_path, from_date, to_date, out_path = sys.argv[1:6]
    old = [e for e in extract_const(html, 'LISTINGS_ALL') if e.get('dealType') in ('전세', '월세')]
    d = json.load(open(deals_path, encoding='utf-8'))
    new = d['jeonse'] + d['wolse']

    # 오늘 파일이 안 올라온 단지는 비교에서 빼야 함(안 그러면 전부 소멸로 잡힘)
    touched = {tuple(t) for t in d.get('touched', [])}
    old = [e for e in old if (e['type'], e['complex']) in touched]

    ob, nb = defaultdict(list), defaultdict(list)
    for e in old: ob[gkey(e)].append(e)
    for e in new: nb[gkey(e)].append(e)

    fresh, gone, changed = [], [], []
    def add_change(o, n):
        if money(o) != money(n):
            c = dict(o)
            c.update({'oldPrice': o['price'], 'newPrice': n['price'],
                      'oldRent': o.get('rent'), 'newRent': n.get('rent')})
            changed.append(c)

    for k in set(ob) | set(nb):
        olds, news = ob.get(k, []), nb.get(k, [])
        o_used, n_used = [False]*len(olds), [False]*len(news)
        # 1차: 매물번호
        for i, o in enumerate(olds):
            ids = set(o.get('listingIds') or [])
            for j, n in enumerate(news):
                if not n_used[j] and ids & set(n.get('listingIds') or []):
                    o_used[i] = n_used[j] = True; add_change(o, n); break
        # 2차: 금액 동일
        for i, o in enumerate(olds):
            if o_used[i]: continue
            for j, n in enumerate(news):
                if not n_used[j] and money(o) == money(n):
                    o_used[i] = n_used[j] = True; break
        ro = [o for i, o in enumerate(olds) if not o_used[i]]
        rn = [n for j, n in enumerate(news) if not n_used[j]]
        # 3차: 개수 같으면 재등록
        if ro and len(ro) == len(rn):
            for o, n in zip(sorted(ro, key=money), sorted(rn, key=money)):
                add_change(o, n)
        else:
            gone += ro; fresh += rn

    out = {'from': from_date, 'to': to_date, 'new': fresh, 'gone': gone, 'changed': changed}
    json.dump(out, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"RENT_DIFF {from_date}→{to_date}: 신규 {len(fresh)} / 소멸 {len(gone)} / 가격변동 {len(changed)}")

if __name__ == '__main__':
    main()

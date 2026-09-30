"""
listing_history.json(화면에서 "이 매물이 언제부터 있었나"를 계산할 때 비동기로 불러오는 가벼운 스냅샷 모음)에
오늘 매매 매물을 추가한다. 같은 날짜가 이미 있으면 덮어씀(당일 재업로드 대응).
오늘 파일이 안 올라온 단지는 직전 날짜 것을 그대로 이어받음(update_archive_and_recompute.py와 같은 규칙).

사용법: python3 update_listing_history.py <listing_history.json> <all_deals.json> <YYYY-MM-DD> <출력.json>
"""
import json, sys

FIELDS = ['gu', 'type', 'complex', 'building', 'floor', 'unitType', 'direction', 'area', 'price']

def main():
    hist_path, deals_path, date, out_path = sys.argv[1:5]
    hist = json.load(open(hist_path, encoding='utf-8'))
    deals = json.load(open(deals_path, encoding='utf-8'))
    sale = deals['sale']
    touched = {(e['type'], e['complex']) for e in sale}
    prev_dates = sorted(d for d in hist if d < date)
    carried = [e for e in hist[prev_dates[-1]] if (e['type'], e['complex']) not in touched] if prev_dates else []
    hist[date] = carried + [{k: e[k] for k in FIELDS} for e in sale]
    hist = dict(sorted(hist.items()))
    json.dump(hist, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"listing_history: {len(hist)}일치, {date} {len(hist[date])}건 (이월 {len(carried)}건)")

if __name__ == '__main__':
    main()

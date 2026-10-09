"""
CRM 단지 브리핑 연동용 공개 파일 listings_today.json 생성 (2026-10-09 추가).
CRM이 하루 한 번 이상 읽어서 "고객에게 담은 매물"의 가격변동·소멸을 판정한다.

형식(CRM과 합의, 리치아파트_답변_브리핑연동_v2.md 2-2):
{ "date": "YYYY-MM-DD",
  "listings": [ { listingKey, listingIds, housingType, complex, dealType, building, floor,
                  unitType, area, price, monthly, direction, checkedDate } ] }
- 값 표기(complex·building·floor·unitType)는 리치아파트 LISTINGS_ALL과 똑같이 둔다
  → 브리핑 화면에서 "담기"로 보내는 값과 같아야 CRM 판정 규칙 2(같은 동·층·타입·가격)가 동작함.
- price: 억(매매가 또는 보증금), monthly: 만원(월세만). 중개사 상호는 넣지 않음.

사용법: python3 build_listings_today.py <오늘_all_deals.json> <YYYY-MM-DD> <출력.json>
"""
import json, sys

def main():
    deals_path, date, out_path = sys.argv[1:4]
    d = json.load(open(deals_path, encoding='utf-8'))
    out = []
    for kind, deal_type in [('sale', '매매'), ('jeonse', '전세'), ('wolse', '월세')]:
        for e in d[kind]:
            ids = sorted(str(i) for i in (e.get('listingIds') or []))
            if not ids:
                continue
            cd = str(e.get('checkedDate') or '')
            out.append({
                'listingKey': '|'.join(ids),
                'listingIds': ids,
                'housingType': e['type'],
                'complex': e['complex'],
                'dealType': deal_type,
                'building': e.get('building'),
                'floor': e.get('floor'),
                'unitType': e.get('unitType'),
                'area': e.get('area'),
                'price': e.get('price'),
                'monthly': e.get('rent') if deal_type == '월세' else None,
                'direction': e.get('direction'),
                'checkedDate': f"{cd[:4]}-{cd[4:6]}-{cd[6:8]}" if len(cd) == 8 else cd,
            })
    json.dump({'date': date, 'listings': out}, open(out_path, 'w', encoding='utf-8'),
              ensure_ascii=False, separators=(',', ':'))
    print(f"listings_today: {date} {len(out)}건 -> {out_path}")

if __name__ == '__main__':
    main()

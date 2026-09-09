"""
집셀 매물 원본 엑셀에서 매매/전세/월세 전부(단, 매매는 dealType 필드 없이 기존 관례 유지) 추출.
전세/월세는 extract_jeonse_wolse.py와 같은 방식, 매매는 LISTINGS_ALL 기존 스키마와 동일하게 맞춤.

사용법: python3 extract_all_deals.py <집셀_파일들_있는_폴더> <출력.json>
"""
import openpyxl, glob, os, re, json, sys
from collections import defaultdict

# 갭투자 판단 키워드 — 대표 매물뿐 아니라 같은 매물을 올린 모든 중복 부동산의 "특징" 문구를 다 훑어서
# 하나라도 걸리면 갭투자로 집계. 세입자가 있어 즉시입주가 어려움을 뜻하는 표현들.
GAP_KEYWORDS = ['세안고', '전세안고', '갭투자', '긴잔금', '무주택', '긴전세', '임차인', '명도불가', '입주불가']

def detect_gap(items):
    for i in items:
        text = (i.get('feature') or '')
        # 원본에 공백이 섞여 들어오는 경우가 있어서(예: "긴 잔 금") 공백 제거하고 검사
        compact = text.replace(' ', '')
        if any(kw in compact for kw in GAP_KEYWORDS):
            return True
    return False

def parse_gu_dong(addr):
    parts = (addr or '').split()
    gu = parts[2] if len(parts) > 2 else '기흥구'
    # "동" 뒤에 지번(숫자)이 더 붙는 경우가 있어서(예: "구갈동 603") 마지막 토큰이 아니라
    # "동"으로 끝나는 토큰을 뒤에서부터 찾아야 함
    dong = ''
    for p in reversed(parts):
        if p.endswith('동'):
            dong = p
            break
    if not dong and parts:
        dong = parts[-1]
    return gu, dong

def process_file(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    info = {}
    if '단지 정보' in wb.sheetnames:
        ws = wb['단지 정보']
        for row in ws.iter_rows(min_row=1, max_row=999999, max_col=5, values_only=True):
            if row and row[0]:
                info[row[0]] = row[1]
    complex_name = info.get('단지명', '?')
    addr = info.get('주소', '')
    gu, dong = parse_gu_dong(addr)
    housing_type = info.get('부동산유형', '아파트')
    if '오피스텔' in os.path.basename(path):
        housing_type = '오피스텔'

    ws = wb['매물 목록'] if '매물 목록' in wb.sheetnames else wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, max_row=999999, values_only=True):
        if not r or r[0] is None:
            continue
        gubun, rep_no, listing_no, deal_type, name, building, ho, floor, unit_type, area, price, rent, direction, checked, confirm, feature, agency, cp, dup, link, memo = (list(r) + [None]*21)[:21]
        rows.append({
            'gubun': gubun, 'rep_no': rep_no, 'listing_no': listing_no, 'dealType': deal_type,
            'building': str(building) if building else None, 'floor': floor, 'unitType': unit_type,
            'area': area, 'price': price, 'rent': rent, 'direction': direction,
            'checked': checked, 'agency': agency, 'feature': feature,
        })

    # 구분/대표매물번호로 그룹핑 (대표 매물 기준으로 중복 부동산 통합)
    groups = defaultdict(list)
    for r in rows:
        key = r['rep_no'] if r['gubun'] == '중복' else r['listing_no']
        groups[key].append(r)

    sale_out, jeonse_out, wolse_out = [], [], []
    for key, items in groups.items():
        rep = next((i for i in items if i['gubun'] == '대표'), items[0])
        dt = rep['dealType']
        agencies = list({i['agency'] for i in items if i['agency']})
        listing_ids = [i['listing_no'] for i in items if i['listing_no']]
        try:
            area_f = float(rep['area'])
        except:
            continue

        def normalize_price(raw):
            # 가격이 이미 억 단위(예: 11.8)인지 원단위(만원, 예: 118000)인지 방어적으로 처리
            if isinstance(raw, (int, float)) and raw > 1000:
                return raw / 10000
            return float(raw)

        if dt == '매매':
            # "대표" 매물 하나의 가격만 쓰면, 같은 매물을 올린 여러 부동산이 서로 다른 호가를 부른 경우
            # (예: 12억/12.5억 혼재) 대표로 뽑힌 쪽이 우연히 더 낮은 값이면 실제보다 낮게 보일 수 있음.
            # 그래서 이 그룹(대표+중복) 안에서 가장 높은 호가를 최종 가격으로 씀.
            try:
                prices = [normalize_price(i['price']) for i in items if i['price'] is not None]
                price_f = max(prices)
            except:
                continue
            sale_out.append({
                'gu': None, 'type': None, 'dong': None, 'complex': complex_name,
                'building': rep['building'], 'floor': str(rep['floor']), 'unitType': str(rep['unitType']),
                'area': area_f, 'price': round(float(price_f), 2), 'direction': rep['direction'],
                'checkedDate': str(rep['checked']), 'gap': detect_gap(items),
                'agencies': agencies, 'listingIds': listing_ids,
            })
        elif dt in ('전세', '월세'):
            try:
                deposits = [normalize_price(i['price']) for i in items if i['price'] is not None]
                deposit = max(deposits)
            except:
                continue
            entry = {
                'gu': None, 'type': None, 'dong': None, 'complex': complex_name,
                'building': rep['building'], 'floor': str(rep['floor']), 'unitType': str(rep['unitType']),
                'area': area_f, 'price': round(deposit, 2), 'direction': rep['direction'],
                'checkedDate': str(rep['checked']), 'dealType': dt,
                'agencies': agencies, 'listingIds': listing_ids,
            }
            if dt == '월세':
                try:
                    # 월세도 마찬가지로 중복 매물 중 최고 월세액을 사용
                    rents = [int(i['rent']) for i in items if i['rent']]
                    entry['rent'] = max(rents) if rents else None
                except:
                    entry['rent'] = None
                wolse_out.append(entry)
            else:
                jeonse_out.append(entry)

    for e in sale_out + jeonse_out + wolse_out:
        e['gu'] = gu; e['dong'] = dong; e['type'] = housing_type
    return sale_out, jeonse_out, wolse_out

def main():
    folder, out_path = sys.argv[1], sys.argv[2]
    files = sorted(glob.glob(os.path.join(folder, '집셀_*.xlsx')))
    all_sale, all_jeonse, all_wolse = [], [], []
    touched = set()
    for f in files:
        s, j, w = process_file(f)
        all_sale += s; all_jeonse += j; all_wolse += w
        if s or j or w:
            complex_name = (s+j+w)[0]['complex']
            housing_type = (s+j+w)[0]['type']
            touched.add((housing_type, complex_name))
    result = {'sale': all_sale, 'jeonse': all_jeonse, 'wolse': all_wolse, 'touched': sorted(list(touched))}
    json.dump(result, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"매매 {len(all_sale)}건 / 전세 {len(all_jeonse)}건 / 월세 {len(all_wolse)}건, 단지 {len(touched)}개 -> {out_path}")

if __name__ == '__main__':
    main()

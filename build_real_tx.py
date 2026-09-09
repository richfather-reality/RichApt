"""
매매 REAL_TX(단지별 전체 실거래 이력, "개별 거래 내역"·"층별 비교"·"월별 추이"에서 쓰임) 재구성 스크립트.

국토부 파일 자체가 "최근 1년치 전체"를 담고 있으므로, 9/7-9/8 diff가 아니라 이 파일 하나로
REAL_TX 전체를 다시 만든다. 신고가(isRecordHigh)는 같은 단지+정확한 전용면적(sizeFloor) 그룹 안에서
계약일 순으로 정렬해 그 시점까지의 최고가를 넘었는지로 계산한다(과거 거래도 전부 포함해 재계산하므로
새 파일을 받을 때마다 이 스크립트를 다시 돌리면 항상 최신 상태로 맞아떨어짐).

사용법: python3 build_real_tx.py <아파트_매매.xlsx> <오피스텔_매매.xlsx> <출력.json>
"""
import openpyxl, json, sys
from collections import defaultdict

def load_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=21, values_only=True):
        if not r or r[1] is None:
            continue
        complex_name = r[5]
        area, ym, day = r[6], r[7], r[8]
        amount_raw, building, floor = r[9], r[10], r[11]
        try:
            amount = float(str(amount_raw).replace(',', '')) / 10000
        except:
            continue
        ym = str(int(ym)); day = str(int(day)).zfill(2)
        date_dot = f"{ym[:4]}.{ym[4:6]}.{day}"
        rows.append({
            'complex': complex_name, 'size': float(str(area).replace(',','')),
            'date': date_dot, 'amount': round(amount, 4), 'floor': str(floor),
            'dong': None if building in (None, '-', '') else str(building),
        })
    return rows

def build_real_tx(rows):
    by_complex = defaultdict(list)
    for r in rows:
        by_complex[r['complex']].append(r)
    out = {}
    for cx, items in by_complex.items():
        items_sorted = sorted(items, key=lambda r: r['date'])
        # 정확한 전용면적(sizeFloor) 그룹별로 최고가를 누적 추적하면서 순서대로 처리
        running_max = {}
        entries = []
        for r in items_sorted:
            size_key = round(r['size'], 2)
            prior_max = running_max.get(size_key)
            is_record = prior_max is None or r['amount'] > prior_max
            delta = None if is_record else round(r['amount'] - prior_max, 2)
            entries.append({
                'date': r['date'], 'size': round(r['size'], 1), 'sizeFloor': size_key,
                'floor': r['floor'], 'amount': round(r['amount'], 2), 'dong': r['dong'] or '-',
                'isRecordHigh': is_record, 'isTopPrice': is_record, 'deltaVsRecentHigh': delta,
            })
            if is_record or size_key not in running_max:
                running_max[size_key] = max(running_max.get(size_key, r['amount']), r['amount'])
        # 최신순으로 보여주는 게 기존 화면 관례와 맞음
        entries.sort(key=lambda e: e['date'], reverse=True)
        out[cx] = entries
    return out

def main():
    apt_path, officetel_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    apt_rows = load_rows(apt_path)
    officetel_rows = load_rows(officetel_path)
    result = {
        '아파트': build_real_tx(apt_rows),
        '오피스텔': build_real_tx(officetel_rows),
    }
    json.dump(result, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"아파트 {len(apt_rows)}건 / 오피스텔 {len(officetel_rows)}건 처리 -> {out_path}")

if __name__ == '__main__':
    main()

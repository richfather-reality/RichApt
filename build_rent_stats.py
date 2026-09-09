"""
전세/월세(전세/월세) 실거래 통계 생성 스크립트
국토부 아파트/오피스텔 전월세 실거래가 엑셀에서 구/동별 집계를 만들어
기존 REAL(매매) 구조와 비슷한 형태의 JSON으로 출력한다.
여러 구(郡/區)에 걸친 파일을 한꺼번에 넘기면 구별로 나눠서 결과를 만든다.

사용법: python3 build_rent_stats.py <출력.json> --apt <아파트_전월세1.xlsx> [<아파트_전월세2.xlsx> ...] --officetel <오피스텔_전월세1.xlsx> [...]
"""
import openpyxl, sys, json, re, argparse
from datetime import datetime, timedelta
from collections import defaultdict


# 원본 파일마다 같은 단지를 다른 이름으로 표기하는 경우가 있어서(집셀 vs 국토부, 혹은 부제 유무 등)
# 여기서 통일함. 안 그러면 실거래분석/시세현황에서 같은 단지가 두 개로 쪼개져서 잡힘.
COMPLEX_NAME_ALIASES = {
    '기흥역지웰푸르지오': '기흥역더퍼스트푸르지오',
    '기흥역롯데캐슬스카이(주상복합)': '기흥역롯데캐슬스카이',
}
def normalize_complex_name(name):
    return COMPLEX_NAME_ALIASES.get(name, name)


def parse_gu(sigungu):
    parts = (sigungu or '').split()
    return parts[2] if len(parts) > 2 else (parts[-1] if parts else '')

def parse_dong(sigungu):
    # "구갈동 603"처럼 동 이름 뒤에 지번이 더 붙는 경우가 있어서 문자열 끝이 아니라
    # "동"으로 끝나는 토큰을 뒤에서부터 찾아야 함 (끝 고정 정규식은 지번이 있으면 매칭 실패함)
    parts = (sigungu or '').split()
    for p in reversed(parts):
        if p.endswith('동'):
            return p
    return sigungu or ''

def load_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=20, values_only=True):
        if not r or r[1] is None:
            continue
        sigungu, complex_name, deal_gubun = r[1], normalize_complex_name(r[5]), r[6]
        area, ym, day = r[7], r[8], r[9]
        deposit_raw, rent_raw, floor = r[10], r[11], r[12]
        if deal_gubun not in ('전세', '월세'):
            continue
        try:
            deposit = float(str(deposit_raw).replace(',', ''))/10000  # 억
        except:
            continue
        try:
            rent = int(str(rent_raw).replace(',', '')) if rent_raw not in (None,'','0',0) else 0
        except:
            rent = 0
        ym = str(int(ym))
        day = str(int(day)).zfill(2)
        date_iso = f"{ym[:4]}-{ym[4:6]}-{day}"
        rows.append({
            'gu': parse_gu(sigungu), 'dong': parse_dong(sigungu), 'complex': complex_name, 'dealType': deal_gubun,
            'area': float(area), 'date': date_iso, 'deposit': round(deposit,4),
            'rent': rent, 'floor': str(floor),
        })
    return rows

def build_stats(rows):
    # dealType -> dong -> list  (이 함수에 넘어오는 rows는 이미 특정 구로 필터링된 상태)
    out = {'전세': {'dong': {}}, '월세': {'dong': {}}}
    all_dates = [r['date'] for r in rows]
    maxdate = max(all_dates) if all_dates else None
    cutoff = (datetime.strptime(maxdate, '%Y-%m-%d') - timedelta(days=30)) if maxdate else None
    maxdate_dt = datetime.strptime(maxdate, '%Y-%m-%d') if maxdate else None

    def dep_unit_price(r): return r['deposit']*10000/r['area']  # 보증금 만원/㎡
    def in_window(items, start, end_exclusive):
        return [r for r in items if start <= datetime.strptime(r['date'],'%Y-%m-%d') < end_exclusive]
    def avg_dep_unit_price(items):
        ups = [dep_unit_price(r) for r in items]
        return sum(ups)/len(ups) if ups else None

    for dt in ('전세', '월세'):
        subset = [r for r in rows if r['dealType'] == dt]

        # 보증금 변동률: 매매와 같은 방식 — 최근 3개월 평균 ㎡당 보증금 vs 6개월 전부터 3개월 전까지 평균 비교
        changePct = None
        if maxdate_dt:
            cur_start = maxdate_dt - timedelta(days=90)
            prev_start = maxdate_dt - timedelta(days=180)
            cur_up = avg_dep_unit_price(in_window(subset, cur_start, maxdate_dt+timedelta(days=1)))
            prev_up = avg_dep_unit_price(in_window(subset, prev_start, cur_start))
            if cur_up and prev_up:
                changePct = round((cur_up-prev_up)/prev_up*100, 2)

        by_dong = defaultdict(list)
        for r in subset:
            by_dong[r['dong']].append(r)
        dong_out = {}
        for dong, items in by_dong.items():
            items_sorted = sorted(items, key=lambda r: r['date'], reverse=True)
            deposits = [i['deposit'] for i in items]
            rents = [i['rent'] for i in items if i['rent']]
            items84 = [i for i in items if 84 <= i['area'] < 85]
            deposits84 = [i['deposit'] for i in items84]
            recent = [{
                'name': i['complex'], 'area': round(i['area'],1), 'floor': i['floor'],
                'date': i['date'], 'dateShort': i['date'][5:].replace('-', '/'),
                'deposit': round(i['deposit'],2), 'rent': i['rent'],
            } for i in items_sorted[:10]]
            # 동 단위 변동률도 구 전체와 같은 방식(최근 3개월 vs 6개월 전 3개월)으로 계산.
            # 이게 빠져있으면 매매 쪽에서 났던 것과 같은 "데이터 부족" 오표시가 생길 수 있음.
            dong_up = avg_dep_unit_price(in_window(items, cur_start, maxdate_dt+timedelta(days=1))) if maxdate_dt else None
            dong_prev_up = avg_dep_unit_price(in_window(items, prev_start, cur_start)) if maxdate_dt else None
            dong_changePct = round((dong_up-dong_prev_up)/dong_prev_up*100, 2) if (dong_up and dong_prev_up) else None
            dong_out[dong] = {
                'count': len(items),
                'avgDeposit': round(sum(deposits)/len(deposits),2) if deposits else None,
                'avgRent': round(sum(rents)/len(rents),1) if rents else None,
                'avgDeposit84': round(sum(deposits84)/len(deposits84),2) if deposits84 else None,
                'count84': len(items84),
                'changePct': dong_changePct,
                'recent': recent,
            }
        total_deposits = [r['deposit'] for r in subset]
        total_rents = [r['rent'] for r in subset if r['rent']]
        total_items84 = [r for r in subset if 84 <= r['area'] < 85]
        total_deposits84 = [r['deposit'] for r in total_items84]
        recent30 = [r for r in subset if cutoff and datetime.strptime(r['date'],'%Y-%m-%d') >= cutoff]

        # 단지별 집계 (실거래분석 탭 전용) — 평형이 단지마다 달라서 전체 평형 평균은 단지간 비교가 불공정함.
        # 그래서 84㎡(국민평형) 거래만 모아서 비교 — 매매 쪽 "84㎡ 평균가"와 같은 기준.
        # 84㎡ 거래가 없는 단지는 이 순위표에서 빠짐(비교 대상 자체가 없으므로).
        items84_all = [r for r in subset if 84 <= r['area'] < 85]
        by_complex_84 = defaultdict(list)
        for r in items84_all:
            by_complex_84[r['complex']].append(r)
        complex_out = []
        for cx, items in by_complex_84.items():
            deposits = [i['deposit'] for i in items]
            rents = [i['rent'] for i in items if i['rent']]
            dongs = {i['dong'] for i in items}
            cx_change = None
            if maxdate_dt:
                cur_up = avg_dep_unit_price(in_window(items, cur_start, maxdate_dt+timedelta(days=1)))
                prev_up = avg_dep_unit_price(in_window(items, prev_start, cur_start))
                if cur_up and prev_up:
                    cx_change = round((cur_up-prev_up)/prev_up*100, 2)
            complex_out.append({
                'complex': cx,
                'dong': sorted(dongs)[0] if len(dongs)==1 else '/'.join(sorted(dongs)),
                'count': len(items),
                'avgDeposit': round(sum(deposits)/len(deposits),2) if deposits else None,
                'minDeposit': round(min(deposits),2) if deposits else None,
                'maxDeposit': round(max(deposits),2) if deposits else None,
                'avgRent': round(sum(rents)/len(rents),1) if rents else None,
                'changePct': cx_change,
            })
        complex_out.sort(key=lambda c: c['count'], reverse=True)


        # 단지+평형별 평균 (내집시세예측용 — 매매의 COMPLEX_STATS와 같은 역할) — 여긴 84㎡로 제한하지 않고 전체 평형 다 포함
        by_complex_all = defaultdict(list)
        for r in subset:
            by_complex_all[r['complex']].append(r)
        by_area_complex = {}
        for cx, items in by_complex_all.items():
            area_buckets = defaultdict(list)
            for i in items:
                area_buckets[str(int(i['area']))].append(i)
            cx_areas = {}
            for area_key, area_items in area_buckets.items():
                deposits = [i['deposit'] for i in area_items]
                rents = [i['rent'] for i in area_items if i['rent']]
                cx_areas[area_key] = {
                    'avgDeposit': round(sum(deposits)/len(deposits),2) if deposits else None,
                    'count': len(deposits),
                    'avgRent': round(sum(rents)/len(rents),1) if rents else None,
                }
            by_area_complex[cx] = cx_areas

        out[dt] = {
            'totalCount': len(subset),
            'recent30Count': len(recent30),
            'avgDeposit': round(sum(total_deposits)/len(total_deposits),2) if total_deposits else None,
            'avgDeposit84': round(sum(total_deposits84)/len(total_deposits84),2) if total_deposits84 else None,
            'count84': len(total_deposits84),
            'avgRent': round(sum(total_rents)/len(total_rents),1) if total_rents else None,
            'changePct': changePct,
            'maxdate': maxdate,
            'dong': dong_out,
            'complexStats': complex_out,
            'byAreaComplex': by_area_complex,
        }
    return out

def main():
    p = argparse.ArgumentParser()
    p.add_argument('out_path')
    p.add_argument('--apt', nargs='+', default=[])
    p.add_argument('--officetel', nargs='+', default=[])
    args = p.parse_args()

    apt_rows = []
    for f in args.apt:
        apt_rows.extend(load_rows(f))
    officetel_rows = []
    for f in args.officetel:
        officetel_rows.extend(load_rows(f))

    all_gus = sorted({r['gu'] for r in apt_rows+officetel_rows})
    result = {}
    for gu in all_gus:
        result[gu] = {
            '아파트': build_stats([r for r in apt_rows if r['gu']==gu]),
            '오피스텔': build_stats([r for r in officetel_rows if r['gu']==gu]),
        }
    with open(args.out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False)
    print('아파트 전월세', len(apt_rows), '건 / 오피스텔 전월세', len(officetel_rows), '건 처리, 구:', all_gus)

if __name__ == '__main__':
    main()

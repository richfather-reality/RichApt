"""
매매 실거래 통계(REAL, COMPLEX_STATS) 생성 스크립트
국토부 아파트/오피스텔 매매 실거래가 엑셀에서 REAL(구/동 단위 요약)과
COMPLEX_STATS(단지+평형[+동] 단위 평균가, 내집시세예측용)를 다시 만든다.

주의: 원래 파이프라인 스크립트(REAL_TX 신고가 배지, EXISTING_MAX)는 유실되어
이 스크립트로는 재현하지 못함 — REAL/COMPLEX_STATS만 새로 만듦.
changePct/change84Pct(변동률)는 "최근 3개월(90일) 평균 ㎡당 단가"와
"6개월 전부터 3개월 전까지의 3개월 평균 ㎡당 단가"를 비교하는 방식으로 계산함.
(30일 vs 이전 60일로 계산했을 때 기존 알려진 값과 부호가 반대로 나왔는데, 이 3개월/3개월
비교 방식으로 하니 기존 값과 거의 일치해서 이 방식으로 확정함 — 완전히 똑같다는 보장은 없음)

사용법:
  python3 build_sale_stats.py <출력_prefix> \
    --apt <아파트_매매1.xlsx> [<아파트_매매2.xlsx> ...] \
    --officetel <오피스텔_매매1.xlsx> [...]
"""
import openpyxl, json, re, argparse
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
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=21, values_only=True):
        if not r or r[1] is None:
            continue
        sigungu, complex_name = r[1], normalize_complex_name(r[5])
        area, ym, day = r[6], r[7], r[8]
        amount_raw, building, floor = r[9], r[10], r[11]
        try:
            price = float(str(amount_raw).replace(',', ''))/10000  # 억
        except:
            continue
        ym = str(int(ym)); day = str(int(day)).zfill(2)
        date_iso = f"{ym[:4]}-{ym[4:6]}-{day}"
        rows.append({
            'gu': parse_gu(sigungu), 'dong': parse_dong(sigungu), 'complex': complex_name,
            'area': float(str(area).replace(',','')), 'date': date_iso, 'price': round(price,4),
            'floor': str(floor), 'building': None if building in (None,'-','') else str(building),
        })
    return rows

def build_real(rows):
    all_dates = [r['date'] for r in rows]
    if not all_dates:
        return {'totalCount':0,'recent30Count':0,'avgPrice':None,'avg84Price':None,'count84':0,
                'changePct':None,'change84Pct':None,'maxdate':None,'dong':{}}
    maxdate = max(all_dates)
    maxdate_dt = datetime.strptime(maxdate, '%Y-%m-%d')
    # 변동률: "최근 3개월 평균 ㎡당 단가" vs "6개월 전부터 3개월 전까지의 3개월 평균 ㎡당 단가" 비교
    # (30일 vs 이전 60일 비교로 했을 때 원래 값과 부호가 반대로 나와서, 3개월/3개월 비교로 바꾸니
    #  기존에 알려진 값(예: 기흥구 84㎡ +3.68%)과 거의 일치해 이 방식으로 확정함)
    cur_start = maxdate_dt - timedelta(days=90)
    prev_start = maxdate_dt - timedelta(days=180)
    prev_end = cur_start
    cutoff30 = maxdate_dt - timedelta(days=30)

    def unit_price(r): return r['price']*10000/r['area']  # 만원/㎡
    def in_window(items, start, end_exclusive):
        return [r for r in items if start <= datetime.strptime(r['date'],'%Y-%m-%d') < end_exclusive]
    def avg_unit_price(items):
        ups = [unit_price(r) for r in items]
        return sum(ups)/len(ups) if ups else None

    cur_all = in_window(rows, cur_start, maxdate_dt + timedelta(days=1))
    prev_all = in_window(rows, prev_start, prev_end)
    cur_up, prev_up = avg_unit_price(cur_all), avg_unit_price(prev_all)
    changePct = round((cur_up-prev_up)/prev_up*100, 2) if (cur_up and prev_up) else None

    items84 = [r for r in rows if 84 <= r['area'] < 85]
    cur_84 = in_window(items84, cur_start, maxdate_dt + timedelta(days=1))
    prev_84 = in_window(items84, prev_start, prev_end)
    cur_up84, prev_up84 = avg_unit_price(cur_84), avg_unit_price(prev_84)
    change84Pct = round((cur_up84-prev_up84)/prev_up84*100, 2) if (cur_up84 and prev_up84) else None

    recent30 = [r for r in rows if datetime.strptime(r['date'],'%Y-%m-%d') >= cutoff30]

    by_dong = defaultdict(list)
    for r in rows: by_dong[r['dong']].append(r)
    dong_out = {}
    for dong, items in by_dong.items():
        items_sorted = sorted(items, key=lambda r: r['date'], reverse=True)
        prices = [i['price'] for i in items]
        items84d = [i for i in items if 84 <= i['area'] < 85]
        prices84 = [i['price'] for i in items84d]
        recent = [{
            'name': i['complex'], 'area': round(i['area'],1), 'floor': i['floor'],
            'date': i['date'], 'dateShort': i['date'][5:].replace('-','/'),
            'price': round(i['price'],2),
        } for i in items_sorted[:10]]
        # 동 단위 변동률도 구 전체와 같은 방식(최근 3개월 vs 6개월 전 3개월)으로 계산.
        # 이게 빠져있으면 화면에 "데이터 부족"으로 잘못 뜸(실제로는 거래가 있어도).
        dong_up = avg_unit_price(in_window(items, cur_start, maxdate_dt + timedelta(days=1)))
        dong_prev_up = avg_unit_price(in_window(items, prev_start, prev_end))
        dong_changePct = round((dong_up-dong_prev_up)/dong_prev_up*100, 2) if (dong_up and dong_prev_up) else None
        dong_up84 = avg_unit_price(in_window(items84d, cur_start, maxdate_dt + timedelta(days=1)))
        dong_prev_up84 = avg_unit_price(in_window(items84d, prev_start, prev_end))
        dong_change84Pct = round((dong_up84-dong_prev_up84)/dong_prev_up84*100, 2) if (dong_up84 and dong_prev_up84) else None
        dong_out[dong] = {
            'count': len(items),
            'avgPrice': round(sum(prices)/len(prices),2) if prices else None,
            'avg84Price': round(sum(prices84)/len(prices84),2) if prices84 else None,
            'count84': len(items84d),
            'changePct': dong_changePct,
            'change84Pct': dong_change84Pct,
            'recent': recent,
        }

    prices_all = [r['price'] for r in rows]
    return {
        'totalCount': len(rows), 'recent30Count': len(recent30),
        'avgPrice': round(sum(prices_all)/len(prices_all),2) if prices_all else None,
        'avg84Price': round(sum(i['price'] for i in items84)/len(items84),2) if items84 else None,
        'count84': len(items84),
        'changePct': changePct, 'change84Pct': change84Pct,
        'maxdate': maxdate, 'dong': dong_out,
    }

def trimmed_avg(prices):
    # 하위 20% 제외 평균 — 시세예측에서 극단적으로 싼 급매/특수거래가 평균을 왜곡하지 않게 함.
    # 표본이 5건 미만이면 20%를 잘라도 통계적으로 의미가 없어서(반올림하면 0건 제외되는 경우가 많음)
    # 그냥 전체 평균을 그대로 씀.
    if len(prices) < 5:
        return round(sum(prices)/len(prices), 2)
    s = sorted(prices)
    cut = int(len(s) * 0.2)
    kept = s[cut:]
    return round(sum(kept)/len(kept), 2)

def build_complex_stats(rows):
    out = {}
    by_complex = defaultdict(list)
    for r in rows: by_complex[r['complex']].append(r)
    for cx, items in by_complex.items():
        area_buckets = defaultdict(list)
        for i in items:
            area_buckets[str(int(i['area']))].append(i)
        cx_out = {}
        for area_key, area_items in area_buckets.items():
            prices = [i['price'] for i in area_items]
            entry = {'_all': {'avg': round(sum(prices)/len(prices),2), 'avgTrimmed': trimmed_avg(prices), 'count': len(prices)}}
            by_building = defaultdict(list)
            for i in area_items:
                if i['building']: by_building[i['building']].append(i['price'])
            for b, bprices in by_building.items():
                entry[b] = {'avg': round(sum(bprices)/len(bprices),2), 'avgTrimmed': trimmed_avg(bprices), 'count': len(bprices)}
            cx_out[area_key] = entry
        out[cx] = cx_out
    return out

def main():
    p = argparse.ArgumentParser()
    p.add_argument('out_prefix')
    p.add_argument('--apt', nargs='+', default=[])
    p.add_argument('--officetel', nargs='+', default=[])
    args = p.parse_args()

    apt_rows = []
    for f in args.apt: apt_rows.extend(load_rows(f))
    officetel_rows = []
    for f in args.officetel: officetel_rows.extend(load_rows(f))

    all_gus = sorted({r['gu'] for r in apt_rows+officetel_rows})

    real_out = {}
    complex_out = {'아파트': {}, '오피스텔': {}}
    for gu in all_gus:
        gu_apt = [r for r in apt_rows if r['gu']==gu]
        gu_off = [r for r in officetel_rows if r['gu']==gu]
        # REAL은 기존에도 아파트 매매 기준 하나만 있었음(오피스텔은 별도 종합 통계 없이 "준비중" 안내로 대체) → 그대로 아파트만
        real_out[gu] = build_real(gu_apt)
        complex_out['아파트'][gu] = build_complex_stats(gu_apt)
        complex_out['오피스텔'][gu] = build_complex_stats(gu_off)

    with open(args.out_prefix+'_real.json','w',encoding='utf-8') as f:
        json.dump(real_out, f, ensure_ascii=False)
    # COMPLEX_STATS 기존 스키마는 gu가 최상위, 그 아래 type이었음 -> 맞춰서 재구성
    complex_final = {}
    for gu in all_gus:
        complex_final[gu] = {
            '아파트': complex_out['아파트'].get(gu, {}),
            '오피스텔': complex_out['오피스텔'].get(gu, {}),
        }
    with open(args.out_prefix+'_complexstats.json','w',encoding='utf-8') as f:
        json.dump(complex_final, f, ensure_ascii=False)

    print('아파트 매매', len(apt_rows), '건 / 오피스텔 매매', len(officetel_rows), '건 처리, 구:', all_gus)

if __name__ == '__main__':
    main()

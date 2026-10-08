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

# 변동률·시세예측 비교 기준 기간(일). 2026-10-02에 90일 → 120일로 변경:
# 90일은 6/30 같은 거래 몰림이 구간 경계를 넘을 때 하루에 5%p 넘게 출렁여서, 표본을 늘려 안정화함.
# 변동률 = 최근 120일 평균 ㎡당 단가 vs 그 이전 120일(240~120일 전) 평균. 시세예측용 단지+평형 평균도 최근 120일만 사용.
WINDOW_DAYS = 120


# 원본 파일마다 같은 단지를 다른 이름으로 표기하는 경우가 있어서(집셀 vs 국토부, 혹은 부제 유무 등)
# 여기서 통일함. 안 그러면 실거래분석/시세현황에서 같은 단지가 두 개로 쪼개져서 잡힘.
COMPLEX_NAME_ALIASES = {
    '기흥역지웰푸르지오': '기흥역더퍼스트푸르지오',
    '기흥역롯데캐슬스카이(주상복합)': '기흥역롯데캐슬스카이',
    # 2026-10-02 추가: 국토부 실거래 단지명이 집셀 매물 단지명과 달라서 실거래가 매물 단지에 안 붙던 7개 단지.
    # (같은 동 + 같은 전용면적 구성으로 동일 단지임을 확인함)
    '강남마을자연앤(6단지)': '강남마을6단지자연앤',
    '코오롱하늘채(5단지)': '강남마을5단지코오롱하늘채',
    '금화마을주공3단지': '금화마을3단지주공그린빌',
    '금화마을주공그린빌6차': '금화마을6단지주공그린빌',
    '금화대우현대1단지': '금화마을대우현대',
    '삼거마을삼성래미안1': '삼거마을삼성래미안1차',
    '성호샤인힐즈아파트': '성호샤인힐즈',
    # 상갈메트로파크는 국토부에 옛 이름(금화마을4단지주공그린빌)으로 등록돼 있음 — 상갈동 463, 금화로58번길 10 주소로 확인
    '금화마을4단지주공그린빌': '상갈메트로파크',
    # 2026-10-08 추가: 국토부가 단지명에 띄어쓰기를 넣어 공개하는 단지들 — 매물(집셀) 이름과 연결이 안 돼서
    # 실거래·신고가·시세예측에 매물 단지로 안 잡혔음(아파트 레이시티, 오피스텔 4개 단지)
    '기흥역 롯데캐슬 레이시티': '기흥역롯데캐슬레이시티',
    '기흥역 더샵': '기흥역더샵',
    '기흥역 센트럴 푸르지오': '기흥역센트럴푸르지오',
    '기흥역 파크 푸르지오': '기흥역파크푸르지오',
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

def load_rows(path, is_officetel=False):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    # 해제(계약취소)된 거래 제외 — 국토부 원본의 '해제사유발생일' 칸에 날짜가 있으면 취소된 계약임.
    # 이걸 안 빼면 취소된 거래가 실거래·신고가·새 실거래 목록에 그대로 잡힘(2026-09-30에 발견, 약 5%).
    # 아파트/오피스텔 파일마다 컬럼 위치가 달라서 13행 헤더에서 이름으로 위치를 찾음.
    header = next(ws.iter_rows(min_row=13, max_row=13, max_col=25, values_only=True), ())
    cancel_idx = list(header).index('해제사유발생일') if '해제사유발생일' in header else None
    rows = []
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=25, values_only=True):
        if not r or r[1] is None:
            continue
        if cancel_idx is not None and r[cancel_idx] not in (None, '', '-'):
            continue
        sigungu, complex_name = r[1], normalize_complex_name(r[5])
        area, ym, day = r[6], r[7], r[8]
        amount_raw = r[9]
        # 국토부 원본 컬럼 순서가 아파트/오피스텔에서 다름(아파트만 "동" 컬럼이 하나 더 있음) —
        # 구분 안 하고 같은 인덱스로 읽으면 오피스텔의 "층" 자리에 "매수자" 값이 잘못 들어감.
        if is_officetel:
            building, floor = None, r[10]
        else:
            building, floor = r[10], r[11]
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
    cur_start = maxdate_dt - timedelta(days=WINDOW_DAYS)
    prev_start = maxdate_dt - timedelta(days=WINDOW_DAYS*2)
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
    # 2026-10-08: 같은 정수 평형(84) 안에 타입(정확한 전용면적 84.5327/84.9545 등)이 여러 개인 단지가 많아서,
    # 각 평형 아래 '_exact'에 국토부 정확한 면적별 최근 WINDOW_DAYS일 거래가격 목록을 같이 넣어둠.
    # 화면(내집시세예측)에서 선택한 타입의 정확한 면적과 맞는 거래가 3건 이상이면 그 타입만으로 평균을 내고,
    # 부족하면 기존처럼 정수 평형 전체 평균을 씀. 최근 1년 안에 거래가 있었던 면적은 0건이어도 키를 남김
    # (그래야 '이 평형에 타입이 여러 개'인지 화면에서 알 수 있음).
    year_sizes = defaultdict(set)
    for r in rows:
        year_sizes[(r['complex'], str(int(r['area'])))].add(round(r['area'], 4))
    # 시세예측용 평균은 최근 WINDOW_DAYS일 거래만 사용(예전엔 1년치 전체를 써서 화면 문구 '최근 3개월'과도 안 맞았음)
    if rows:
        maxdate_dt = datetime.strptime(max(r['date'] for r in rows), '%Y-%m-%d')
        cutoff = maxdate_dt - timedelta(days=WINDOW_DAYS)
        rows = [r for r in rows if datetime.strptime(r['date'], '%Y-%m-%d') > cutoff]
    out = {}
    by_complex = defaultdict(list)
    for r in rows: by_complex[r['complex']].append(r)
    for cx, items in by_complex.items():
        # 정수 버림(int()) 기준으로 묶음 — 84.0~84.9999는 항상 "84"로 통일되고, 113.94짜리와
        # 114.03짜리처럼 정수 자체가 다르면(=서로 다른 타입일 가능성) 따로 집계됨.
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
    for (cx, area_key), sizes in year_sizes.items():
        if len(sizes) < 2:
            continue  # 타입이 하나뿐인 평형은 정수 평균 = 타입 평균이라 따로 둘 필요 없음
        recent = defaultdict(list)
        for r in by_complex.get(cx, []):
            if str(int(r['area'])) == area_key:
                recent[round(r['area'], 4)].append(round(r['price'], 2))
        if area_key not in out.get(cx, {}):
            continue  # 최근 WINDOW_DAYS일에 이 평형 거래가 아예 없으면 기존처럼 통계 없음으로 둠
        out[cx][area_key]['_exact'] = {f"{s:.4f}": {'prices': sorted(recent.get(s, []))} for s in sorted(sizes)}
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
    for f in args.officetel: officetel_rows.extend(load_rows(f, is_officetel=True))

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

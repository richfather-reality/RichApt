"""
전월세 "오늘 새로 올라온 실거래"(NEW_RENT_TX) 생성 스크립트.
직전 업로드한 국토부 전월세 파일과 최신 파일을 비교해서, 직전엔 없었는데 최신 파일에 새로 나타난 행만 뽑는다.
(계약일이 오늘이라는 뜻이 아니라 "국토부 시스템에 새로 공개된 거래"라는 뜻 — 매매 NEW_TX와 같은 방식)

어제 파일이 없으면 --apt-old/--officetel-old를 빼고 돌리면 됨 → hasPrev=false, new=[] 로 출력
(화면에서는 "비교할 이전 자료 없음"으로 처리됨).

사용법:
  python3 build_new_rent_tx.py <출력.json> \
    --apt-new <아파트_전월세_오늘.xlsx> [--apt-old <아파트_전월세_어제.xlsx>] \
    --officetel-new <오피스텔_전월세_오늘.xlsx> [--officetel-old <오피스텔_전월세_어제.xlsx>]
"""
import openpyxl, json, argparse
from collections import Counter

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
    parts = (sigungu or '').split()
    for p in reversed(parts):
        if p.endswith('동'):
            return p
    return ''

def load_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=20, values_only=True):
        if not r or r[1] is None:
            continue
        sigungu, complex_raw, gubun = r[1], r[5], r[6]
        area, ym, day, deposit_raw, rent_raw, floor = r[7], r[8], r[9], r[10], r[11], r[12]
        if gubun not in ('전세', '월세'):
            continue
        try:
            deposit = float(str(deposit_raw).replace(',', '')) / 10000
        except:
            continue
        try:
            rent = int(str(rent_raw).replace(',', '')) if rent_raw not in (None, '', '0', 0) else 0
        except:
            rent = 0
        ym = str(int(ym)); day = str(int(day)).zfill(2)
        # 매칭 키: 원본 문자열 그대로(단지·면적·계약일·보증금·월세·층·구분). 같은 날 같은 조건 거래가
        # 여러 건일 수 있어서 set이 아니라 개수(Counter)로 비교함.
        key = (complex_raw, str(area), ym, day, str(deposit_raw), str(rent_raw), str(floor), gubun)
        rows.append({
            'key': key, 'complex': normalize_complex_name(complex_raw),
            'gu': parse_gu(sigungu), 'dong': parse_dong(sigungu),
            'area': round(float(str(area).replace(',', '')), 1),
            'date': f"{ym[:4]}-{ym[4:6]}-{day}", 'dateShort': f"{ym[4:6]}/{day}",
            'deposit': round(deposit, 2), 'rent': rent, 'floor': str(floor), 'dealType': gubun,
        })
    return rows

def diff_new(old_rows, new_rows):
    remaining = Counter(r['key'] for r in old_rows)
    out = []
    for r in new_rows:
        if remaining[r['key']] > 0:
            remaining[r['key']] -= 1
        else:
            out.append(r)
    return out

def main():
    p = argparse.ArgumentParser()
    p.add_argument('out_path')
    p.add_argument('--apt-new', required=True)
    p.add_argument('--apt-old')
    p.add_argument('--officetel-new')
    p.add_argument('--officetel-old')
    args = p.parse_args()

    result = {}
    for htype, new_path, old_path in [('아파트', args.apt_new, args.apt_old),
                                       ('오피스텔', args.officetel_new, args.officetel_old)]:
        if not new_path:
            continue
        new_rows = load_rows(new_path)
        has_prev = bool(old_path)
        fresh = diff_new(load_rows(old_path), new_rows) if has_prev else []
        gus = {r['gu'] for r in new_rows}
        for gu in gus:
            items = [{k: v for k, v in r.items() if k != 'key'} for r in fresh if r['gu'] == gu]
            items.sort(key=lambda x: x['date'], reverse=True)
            result.setdefault(gu, {})[htype] = {'hasPrev': has_prev, 'new': items}

    json.dump(result, open(args.out_path, 'w', encoding='utf-8'), ensure_ascii=False)
    for gu, v in result.items():
        print(gu, {t: len(x['new']) for t, x in v.items()})

if __name__ == '__main__':
    main()

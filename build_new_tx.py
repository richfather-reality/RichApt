"""
매매 "오늘 새로 올라온 실거래"(NEW_TX) 생성 스크립트.
직전 업로드 파일과 최신 파일을 비교해서, 직전엔 없었는데 최신 파일에 새로 나타난 행만 뽑는다.
(계약일이 오늘이라는 뜻이 아니라 "국토부 시스템에 새로 공개된 거래"라는 뜻 — 전월세 NEW_RENT_TX와 같은 방식)

기존 REAL_TX(과거 전체 이력)를 참고해서 이 신규 거래가 그 단지+정확한 평형(sizeFloor 기준) 역대
최고가인지(isRecordHigh)도 같이 계산한다.

사용법:
  python3 build_new_tx.py <REAL_TX_있는_index.html> <출력.json> \
    --apt-old <아파트_매매_어제.xlsx> --apt-new <아파트_매매_오늘.xlsx> \
    --officetel-old <오피스텔_매매_어제.xlsx> --officetel-new <오피스텔_매매_오늘.xlsx>
"""
import openpyxl, json, re, argparse


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
            amount = float(str(amount_raw).replace(',', ''))/10000  # 억
        except:
            continue
        ym = str(int(ym)); day = str(int(day)).zfill(2)
        date_dot = f"{ym[:4]}.{ym[4:6]}.{day}"  # 기존 REAL_TX 표기(점 구분)와 맞춤
        # 매칭 키에서 "동" 정보는 뺌 — 국토부가 처음엔 동을 "-"(공란)로 공개했다가 나중에
        # 실제 동 번호로 보완해서 재공개하는 경우가 있는데, 동을 키에 넣으면 같은 거래를
        # "동이 바뀐 새 거래"로 잘못 인식해버림. 단지+평형+계약일+금액이면 사실상 유일하게 식별됨.
        key = (complex_name, str(area), ym, day, str(amount_raw))
        rows.append({
            'key': key, 'gu': parse_gu(sigungu), 'complex': complex_name,
            'area': float(str(area).replace(',','')), 'date': date_dot, 'amount': round(amount,4),
            'floor': str(floor), 'building': None if building in (None,'-','') else str(building),
        })
    return rows

def extract_real_tx(html_path):
    content = open(html_path, encoding='utf-8').read()
    idx = content.find('const REAL_TX = {')
    start = idx + len('const REAL_TX = ')
    depth=0; in_str=False; esc=False; s=None
    for i in range(start, len(content)):
        c = content[i]
        if s is None:
            if c in '{[': s=i; depth=1; continue
            else: continue
        if in_str:
            if esc: esc=False
            elif c=='\\': esc=True
            elif c=='"': in_str=False
            continue
        else:
            if c=='"': in_str=True
            elif c in '{[': depth+=1
            elif c in '}]':
                depth-=1
                if depth==0: return json.loads(content[s:i+1])
    return {}

def diff_new(old_rows, new_rows):
    old_keys = {r['key'] for r in old_rows}
    return [r for r in new_rows if r['key'] not in old_keys]

def build_new_tx_for_type(new_entries, real_tx_bucket):
    out = {}
    for r in new_entries:
        cx = r['complex']
        existing = real_tx_bucket.get(cx, [])
        # 2026-10-08: 정확한 전용면적(소수 4자리) 기준으로 같은 타입끼리만 비교 (build_real_tx.py와 같은 기준)
        area_key = round(r['area'], 4)
        same_size = [t for t in existing if round(float(t['sizeFloor']), 4) == area_key]
        prior_max = max((t['amount'] for t in same_size), default=None)
        is_record = prior_max is None or r['amount'] > prior_max
        # 신고가면 기존 최고가보다 얼마나 올랐는지(+), 아니면 기존 최고가 대비 얼마나 낮은지(-) — 같은 정확한 전용면적 기준.
        # (이 면적 첫 거래라 비교 대상이 없으면 None)
        delta = None if prior_max is None else round(r['amount'] - prior_max, 2)
        entry = {
            'date': r['date'], 'size': round(r['area'],2), 'sizeFloor': area_key,
            'floor': r['floor'], 'amount': round(r['amount'],2),
            'dong': r['building'] if r['building'] else '-',
            'isRecordHigh': is_record, 'isTopPrice': is_record,
            'deltaVsRecentHigh': delta,
        }
        out.setdefault(cx, []).append(entry)
    for cx in out:
        out[cx].sort(key=lambda t: t['date'])
    return out

def main():
    p = argparse.ArgumentParser()
    p.add_argument('html_path')
    p.add_argument('out_path')
    p.add_argument('--apt-old', required=True)
    p.add_argument('--apt-new', required=True)
    p.add_argument('--officetel-old')
    p.add_argument('--officetel-new')
    args = p.parse_args()

    real_tx = extract_real_tx(args.html_path)

    apt_old = load_rows(args.apt_old)
    apt_new = load_rows(args.apt_new)
    apt_diff = diff_new(apt_old, apt_new)
    apt_new_tx = build_new_tx_for_type(apt_diff, real_tx.get('아파트', {}))

    result = {'아파트': apt_new_tx, '오피스텔': {}}

    if args.officetel_old and args.officetel_new:
        off_old = load_rows(args.officetel_old, is_officetel=True)
        off_new = load_rows(args.officetel_new, is_officetel=True)
        off_diff = diff_new(off_old, off_new)
        result['오피스텔'] = build_new_tx_for_type(off_diff, real_tx.get('오피스텔', {}))

    with open(args.out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False)

    apt_count = sum(len(v) for v in result['아파트'].values())
    off_count = sum(len(v) for v in result['오피스텔'].values())
    print(f'아파트 신규 {apt_count}건, 오피스텔 신규 {off_count}건 -> {args.out_path}')

if __name__ == '__main__':
    main()

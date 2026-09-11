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
}
def normalize_complex_name(name):
    return COMPLEX_NAME_ALIASES.get(name, name)


def parse_gu(sigungu):
    parts = (sigungu or '').split()
    return parts[2] if len(parts) > 2 else (parts[-1] if parts else '')

def load_rows(path, is_officetel=False):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=14, max_row=999999, max_col=21, values_only=True):
        if not r or r[1] is None:
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
        area_floor = int(r['area'])
        same_size = [t for t in existing if int(t['sizeFloor']) == area_floor]
        prior_max = max((t['amount'] for t in same_size), default=None)
        is_record = prior_max is None or r['amount'] > prior_max
        delta = None if is_record else round(r['amount'] - prior_max, 2)
        entry = {
            'date': r['date'], 'size': round(r['area'],2), 'sizeFloor': r['area'],
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

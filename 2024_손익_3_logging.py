import pandas as pd
import xlwings as xw
import re, math
import numpy as np
import logging
from datetime import datetime

# ===== 로깅 설정 =====
def setup_logging():
    """로깅 설정"""
    # 로그 파일명에 현재 시각 포함
    log_filename = f"atm_processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    # 로거 설정
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()  # 콘솔에도 출력
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("=== ATM 데이터 처리 시작 ===")
    logger.info(f"로그 파일: {log_filename}")
    return logger

# ===== 설정 =====
FILE_PATH = r"Y:\\ATM사업실\\고중걸\\0_종합파일\\종합_2024_통합데이터.xlsx"
SHEET_IN = "2024"
SHEET_OUT = "요약"

REQUIRED_COLS = [
    '기번1','대분류1','중분류1','일건수','매출합','공헌이익','영업이익','CD/ATM'
]

NUMERIC_COLS = [
    '일건수','매출합','공헌이익','영업이익','수수료매출','임차료','현송비',
    '중계/명세','유지보수','회선비','장애대처/경비','수선공사/청소',
    '감가상각비','기타고정비','인건비/경비','준고정비'
]

MAJOR_BRANDS = ['CU','세븐일레븐','이마트24']

ONLY_SUBTOTAL_GROUPS = {
    '브랜드','브랜드제휴','일반장소','일반(기타기관+개별장소)','점주대행'
}

GRADE_RULES = [
    (20.0,'A등급'), (10.0,'B등급'), (7.0,'C등급'),
    (5.0,'D등급'), (2.0,'F등급'), (1.0,'X등급'), (0.0,'Y등급')
]

MANUFACTURER_PATTERNS = {
    r'푸른':'푸른', r'네오텍':'네오텍', r'효성':'효성',
    r'청호':'청호', r'에이텍|구\s*LG|LG':'에이텍'
}

# ===== 정확한 데이터 시작 위치 =====
DATA_START_CELLS = {
    'final': 'A7',     # 채널별 운영현황
    'final2': 'D22',   # 편의점 채널 시장현황
    'final3': 'C32',   # 등급별 운영현황
    'final4': 'C43',   # 기기 운영현황
    'final5': 'B54',   # 장애대응 현황
    'final6': 'C61'    # 채널별 원가현황
}

# 6번 표는 특별히 범위 지정 (C61:G68, 점주대행 컬럼 H는 건드리지 않음)
FINAL6_END_CELL = 'G68'

# ===== 유틸 함수 =====
def map_group_for_summary(row):
    """1번 표 전용: 점주대행도 그룹에 포함"""
    if row.get('중분류1') in MAJOR_BRANDS:
        return '편의점'
    dv = row.get('대분류1')
    if dv in ['브랜드', '브랜드제휴']:
        return '브랜드'
    if dv in ['일반장소', '일반(기타기관+개별장소)']:
        return '일반장소'
    if dv == '점주대행':
        return '점주대행'
    return None

def map_group_for_financial(row):
    """6번 표 전용: 점주대행 제외"""
    if row.get('중분류1') in MAJOR_BRANDS:
        return '편의점'
    dv = row.get('대분류1')
    if dv in ['브랜드', '브랜드제휴']:
        return '브랜드'
    if dv in ['일반장소', '일반(기타기관+개별장소)']:
        return '일반장소'
    # 점주대행 없음
    return None

def pct(n, d):
    return f"{(n/d*100):.1f}%" if d not in (0,None) and d!=0 else "0%"

def clean_numeric(s):
    return pd.to_numeric(
        s.astype(str)
         .str.replace(r'\s','',regex=True)
         .str.replace(',','',regex=False)
         .str.replace(r'\(([^)]+)\)',r'-\1',regex=True),
        errors='coerce'
    )

def norm_maker(x):
    if pd.isna(x): return '기타'
    s = str(x).strip()
    for pat, lab in MANUFACTURER_PATTERNS.items():
        if re.search(pat, s):
            return lab
    return '기타'

def get_grade(cnt):
    for th, g in GRADE_RULES:
        if cnt >= th: return g
    return 'Y등급'

def safe_div(a, b):
    return math.nan if pd.isna(a) or pd.isna(b) or b==0 else a/b

# ===== 데이터 로드 =====
def load_and_clean_data(ws, logger):
    logger.info("📊 데이터 로드 및 전처리 시작")
    
    df = ws.used_range.options(pd.DataFrame, header=1, index=False).value
    logger.info(f"원본 데이터 크기: {df.shape}")
    
    df.columns = df.columns.str.strip()
    logger.info(f"컬럼 목록: {list(df.columns)}")

    miss = [c for c in REQUIRED_COLS if c not in df.columns]
    if miss:
        logger.error(f"❌ 필수 열 누락: {miss}")
        raise ValueError(f"❌ 필수 열 누락: {miss}")
    
    logger.info("✅ 필수 컬럼 확인 완료")

    # 숫자형 컬럼 처리
    logger.info("🔢 숫자형 컬럼 전처리 시작")
    for c in NUMERIC_COLS:
        if c not in df.columns:
            df[c] = pd.NA
            logger.warning(f"⚠️ 컬럼 '{c}' 없음 - 기본값으로 설정")
        else:
            before_count = df[c].notna().sum()
            df[c] = clean_numeric(df[c])
            after_count = df[c].notna().sum()
            logger.info(f"컬럼 '{c}': {before_count} → {after_count} 유효값")

    # 문자형 컬럼 처리
    for c in ['CD/ATM','대분류1','중분류1','장애대응']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    
    logger.info(f"✅ 데이터 전처리 완료 - 최종 크기: {df.shape}")
    return df


# ===== 1. 채널별 운영현황 =====
def make_summary_data(df, logger):
    logger.info("📈 1번 표 생성: 채널별 운영현황")
    
    # 월별 평균 계산
    monthly_avg = df.groupby('기번1',dropna=False)['일건수'].mean().reset_index(name='월평균')
    logger.info(f"기기별 월평균 계산 완료: {len(monthly_avg)}개 기기")
    
    device_info = df.drop_duplicates('기번1')[['기번1','대분류1','중분류1']]
    monthly_avg = monthly_avg.merge(device_info, on='기번1', how='left')

    avg_result = monthly_avg.groupby(['대분류1','중분류1'])['월평균'].mean().reset_index()
    avg_result['평균건수'] = avg_result['월평균'].round(1)

    pivot = df.groupby(['대분류1','중분류1']).agg(
        대수=('기번1','nunique'),
        매출액=('매출합','sum'),
        공헌이익=('공헌이익','sum'),
        영업이익=('영업이익','sum')
    ).reset_index()

    logger.info(f"그룹별 집계 완료: {len(pivot)}개 그룹")

    tot_units = pivot['대수'].sum()
    tot_sales = pivot['매출액'].sum()
    tot_profit = pivot['영업이익'].sum()
    
    logger.info(f"전체 집계 - 대수: {tot_units}, 매출: {tot_sales:,}, 영업이익: {tot_profit:,}")

    pivot['대수비중'] = pivot['대수'].apply(lambda v:pct(v, tot_units))
    pivot['매출액비중'] = pivot['매출액'].apply(lambda v:pct(v, tot_sales))
    pivot['영업이익비중'] = pivot['영업이익'].apply(lambda v:pct(v, tot_profit))

    pivot = pivot.merge(avg_result[['대분류1','중분류1','평균건수']],
                        on=['대분류1','중분류1'], how='left')

    data_rows = []
    for group, grp in pivot.groupby('대분류1',sort=False):
        logger.info(f"그룹 '{group}' 처리 중 - {len(grp)}개 하위 분류")
        
        if group in ONLY_SUBTOTAL_GROUPS:
            subtotal = [
                group, '소계',
                grp['대수'].sum(),
                pct(grp['대수'].sum(), tot_units),
                round(monthly_avg.loc[monthly_avg['대분류1']==group,'월평균'].mean(),1),
                grp['매출액'].sum(),
                pct(grp['매출액'].sum(), tot_sales),
                grp['공헌이익'].sum(),
                grp['영업이익'].sum(),
                pct(grp['영업이익'].sum(), tot_profit)
            ]
            data_rows.append(subtotal)
        else:
            for _, row in grp.iterrows():
                data_rows.append([
                    row['대분류1'], row['중분류1'], row['대수'], row['대수비중'],
                    row['평균건수'], row['매출액'], row['매출액비중'],
                    row['공헌이익'], row['영업이익'], row['영업이익비중']
                ])
            subtotal = [
                group, '소계',
                grp['대수'].sum(),
                pct(grp['대수'].sum(), tot_units),
                round(monthly_avg.loc[monthly_avg['대분류1']==group,'월평균'].mean(),1),
                grp['매출액'].sum(),
                pct(grp['매출액'].sum(), tot_sales),
                grp['공헌이익'].sum(),
                grp['영업이익'].sum(),
                pct(grp['영업이익'].sum(), tot_profit)
            ]
            data_rows.append(subtotal)
    
    logger.info(f"✅ 1번 표 생성 완료: {len(data_rows)}행")
    return data_rows


# ===== 2. 편의점 채널 시장현황 =====
def make_convenience_data(df, logger):
    logger.info("🏪 2번 표 생성: 편의점 채널 시장현황")
    
    df_major = df[df['중분류1'].isin(MAJOR_BRANDS)].copy()
    logger.info(f"주요 편의점 브랜드 데이터: {len(df_major)}행")
    
    pvt = df_major.groupby(['중분류1','CD/ATM'])['기번1'].nunique().unstack(fill_value=0)

    for c in ['ATM','CD']:
        if c not in pvt.columns:
            pvt[c] = 0

    pvt['합계'] = pvt['ATM'] + pvt['CD']
    total_units = pvt['합계'].sum()
    pvt['비중'] = pvt['합계'].apply(lambda v: pct(v, total_units))

    logger.info(f"편의점별 집계 완료 - 전체 대수: {total_units}")

    data_rows = []
    for brand in MAJOR_BRANDS:
        if brand in pvt.index:
            row = pvt.loc[brand]
            data_rows.append([row['ATM'], row['CD'], row['합계'], row['비중']])
            logger.info(f"{brand}: ATM {row['ATM']}, CD {row['CD']}, 합계 {row['합계']}")
        else:
            data_rows.append([0, 0, 0, "0%"])
            logger.warning(f"{brand}: 데이터 없음")
    
    logger.info(f"✅ 2번 표 생성 완료: {len(data_rows)}행")
    return data_rows


# ===== 3. 등급별 운영현황 =====
def make_grade_data(df, logger):
    logger.info("📊 3번 표 생성: 등급별 운영현황")
    
    if '등급' not in df.columns:
        logger.info("등급 컬럼 없음 - 일건수 기준으로 등급 생성")
        grade_avg = df.groupby('기번1')['일건수'].sum().reset_index(name='연간')
        grade_avg['평균건수'] = (grade_avg['연간']/12).round(2)
        grade_avg['등급'] = grade_avg['평균건수'].apply(get_grade)
        df = df.merge(grade_avg[['기번1','등급']], on='기번1', how='left')
        
        grade_counts = grade_avg['등급'].value_counts()
        logger.info(f"등급별 기기 수: {grade_counts.to_dict()}")

    grades = [g for _,g in GRADE_RULES]

    pvt = df[['기번1','CD/ATM','등급']].drop_duplicates().pivot_table(
        index='CD/ATM', columns='등급', values='기번1',
        aggfunc='nunique', fill_value=0
    )

    for g in grades:
        if g not in pvt.columns:
            pvt[g] = 0

    pvt = pvt[grades]
    pvt.loc['합계'] = pvt.sum()

    atm_sum = pvt.loc['ATM'].sum() if 'ATM' in pvt.index else 0
    cd_sum  = pvt.loc['CD'].sum()  if 'CD'  in pvt.index else 0
    total_sum = pvt.loc['합계'].sum()

    logger.info(f"등급별 집계 완료 - ATM: {atm_sum}, CD: {cd_sum}, 전체: {total_sum}")

    data_rows = []
    # ATM
    atm_row = list(pvt.loc['ATM']) if 'ATM' in pvt.index else [0]*len(grades)
    atm_row.append(atm_sum)
    data_rows.append(atm_row)
    data_rows.append([pct(pvt.loc['ATM',g] if 'ATM' in pvt.index else 0, atm_sum) for g in grades] + ['100%'])
    # CD
    cd_row = list(pvt.loc['CD']) if 'CD' in pvt.index else [0]*len(grades)
    cd_row.append(cd_sum)
    data_rows.append(cd_row)
    data_rows.append([pct(pvt.loc['CD',g] if 'CD' in pvt.index else 0, cd_sum) for g in grades] + ['100%'])
    # 합계
    total_row = list(pvt.loc['합계'])
    total_row.append(total_sum)
    data_rows.append(total_row)
    data_rows.append([pct(pvt.loc['합계',g], total_sum) for g in grades] + ['100%'])

    logger.info(f"✅ 3번 표 생성 완료: {len(data_rows)}행")
    return data_rows


# ===== 4. 기기 운영현황 (설치장소 + 제조사) =====
def make_location_maker_data(df, logger):
    logger.info("🏭 4번 표 생성: 기기 운영현황")
    
    final4_cols=['점내','점외','네오텍','푸른','효성','에이텍','청호']
    if ('점내,외' not in df.columns) or ('제조사' not in df.columns):
        logger.warning("점내,외 또는 제조사 컬럼 없음")
        return []

    dev = df[['기번1','CD/ATM','점내,외','제조사']].drop_duplicates().copy()
    dev['제조사_정규화'] = dev['제조사'].apply(norm_maker)
    
    logger.info(f"기기 정보 추출 완료: {len(dev)}개 기기")

    loc_counts = dev.pivot_table(index='CD/ATM',columns='점내,외',values='기번1',
                                 aggfunc='nunique',fill_value=0)
    for c in ['점내','점외']:
        if c not in loc_counts.columns:
            loc_counts[c]=0
    loc_counts=loc_counts[['점내','점외']]

    mkr_counts=dev.pivot_table(index='CD/ATM',columns='제조사_정규화',values='기번1',
                               aggfunc='nunique',fill_value=0)
    for m in ['네오텍','푸른','효성','에이텍','청호']:
        if m not in mkr_counts.columns:
            mkr_counts[m]=0
    mkr_counts=mkr_counts[['네오텍','푸른','효성','에이텍','청호']]

    counts=pd.concat([loc_counts,mkr_counts],axis=1)
    for r in ['ATM','CD']:
        if r not in counts.index:
            counts.loc[r]=0
    counts=counts.loc[['ATM','CD'],final4_cols].astype(int)
    counts['합계']=counts['점내']+counts['점외']

    logger.info(f"장소별/제조사별 집계 완료")
    logger.info(f"ATM: {counts.loc['ATM'].to_dict()}")
    logger.info(f"CD: {counts.loc['CD'].to_dict()}")

    def row_pct(row):
        denom=row['합계']
        return [pct(row[c],denom) if denom>0 else '0%' for c in final4_cols] + ['100%']

    data_rows = []
    # ATM
    atm_row = list(counts.loc['ATM'])
    data_rows.append(atm_row)
    data_rows.append(row_pct(counts.loc['ATM']))
    # CD
    cd_row = list(counts.loc['CD'])
    data_rows.append(cd_row)
    data_rows.append(row_pct(counts.loc['CD']))
    # 합계
    total_counts = counts.loc['ATM'] + counts.loc['CD']
    total_row = list(total_counts)
    data_rows.append(total_row)
    total_sum = total_counts['합계']
    total_pct = [pct(total_counts[c],total_sum) if total_sum>0 else '0%' for c in final4_cols] + ['100%']
    data_rows.append(total_pct)

    logger.info(f"✅ 4번 표 생성 완료: {len(data_rows)}행")
    return data_rows


# ===== 5. 장애 대응 현황 =====
def make_issue_data(df, logger):
    logger.info("🔧 5번 표 생성: 장애 대응 현황")
    
    mapping = {
        "전체외주":"전체외주",
        "주간제조사+야간경비사(매일)":"주간제조사\n야간경비사",
        "주간제조사+야간경비사(주말24)":"주간제조사\n야간경비사\n+(주말24시)",
        "주간경비사(매일)+야간미대응":"주간경비사\n야간미대응",
        "주간경비사(주말24)+야간미대응":"주간경비사\n야간미대응\n+(주말24시)",
        "주간CS+야간미대응":"주간CS\n+야간미대응",
        "주간제조사+야간미대응":"주간제조사\n야간미대응"
    }
    if '장애대응' not in df.columns:
        logger.warning("장애대응 컬럼 없음")
        return []
    
    tmp = df[['기번1','장애대응']].copy()
    tmp['장애대응'] = tmp['장애대응'].astype(str).str.strip()
    counts, total = [], 0
    
    for key, label in mapping.items():
        cnt = tmp.loc[tmp['장애대응'] == key, '기번1'].nunique()
        counts.append(cnt)
        total += cnt
        logger.info(f"장애대응 '{key}': {cnt}대")
    
    counts.append(total)
    pcts = [pct(c, total) for c in counts[:-1]] + ['100%']
    
    logger.info(f"✅ 5번 표 생성 완료: 전체 {total}대")
    return [counts, pcts]


# ===== 6. 채널별 원가현황 =====   
    """6번째 표: 채널별 원가현황 (C61부터) - 월별 기준 평균값 (대수/건당수수료 제외)
    점주대행 제외, 5개 컬럼만 출력
    """
    logger.info("💰 6번 표 생성: 채널별 원가현황 (월별 평균 계산, 점주대행 제외)")
    
    groups = ['편의점', '브랜드', '일반장소']
    devices = ['ATM', 'CD']
    months = sorted(df['month'].dropna().unique())
    
    logger.info(f"분석 대상 월: {len(months)}개월")

    monthly_tables = []

    for m in months:
        cur_month = pd.to_datetime(m)
        logger.info(f"월별 데이터 처리 중: {cur_month.strftime('%Y-%m')}")

        df_active = df[(df['month'] == cur_month) & (df['최초설치일'] <= cur_month)].copy()
        if df_active.empty:
            logger.warning(f"{cur_month.strftime('%Y-%m')}: 활성 데이터 없음")
            continue

        df_active['그룹'] = df_active.apply(lambda row: (
            '편의점' if row['중분류1'] in MAJOR_BRANDS else
            '브랜드' if row['대분류1'] in ['브랜드', '브랜드제휴'] else
            '일반장소' if row['대분류1'] in ['일반장소', '일반(기타기관+개별장소)'] else
            None  # 점주대행 제외
        ), axis=1)
        df_active = df_active[df_active['그룹'].notna() & df_active['CD/ATM'].isin(['ATM', 'CD'])]

        logger.info(f"{cur_month.strftime('%Y-%m')}: 유효 데이터 {len(df_active)}행")

        rows = []
        for group in groups:
            if group == '브랜드':
                # ATM + CD 합산
                sub_atm = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == 'ATM')]
                sub_cd  = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == 'CD')]
                sub = pd.concat([sub_atm, sub_cd])
                device_count = sub['기번1'].nunique()
                if device_count == 0:
                    rows.append([np.nan] * 5)  # 5개 컬럼만
                    continue

                sums = sub[['매출합','영업이익','임차료','현송비','중계/명세','유지보수','회선비',
                            '장애대처/경비','수선공사/청소','감가상각비','기타고정비','인건비/경비','준고정비']].sum(numeric_only=True)

                var_cost = sums[['임차료','현송비','중계/명세','유지보수','회선비','장애대처/경비','수선공사/청소']].sum()
                fix_cost = sums[['감가상각비','기타고정비','인건비/경비','준고정비']].sum()
                total_cost = var_cost + fix_cost
                vals = [
                    safe_div(sums['매출합'], device_count),
                    safe_div(total_cost, device_count),
                    safe_div(var_cost, device_count),
                    safe_div(fix_cost, device_count),
                    safe_div(sums['영업이익'], device_count)
                ]
                rows.append(vals)
            else:
                for device in devices:
                    sub = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == device)]
                    device_count = sub['기번1'].nunique()
                    if device_count == 0:
                        rows.append([np.nan] * 5)  # 5개 컬럼만
                        continue

                    sums = sub[['매출합','영업이익','임차료','현송비','중계/명세','유지보수','회선비',
                                '장애대처/경비','수선공사/청소','감가상각비','기타고정비','인건비/경비','준고정비']].sum(numeric_only=True)

                    var_cost = sums[['임차료','현송비','중계/명세','유지보수','회선비','장애대처/경비','수선공사/청소']].sum()
                    fix_cost = sums[['감가상각비','기타고정비','인건비/경비','준고정비']].sum()
                    total_cost = var_cost + fix_cost

                    vals = [
                        safe_div(sums['매출합'], device_count),
                        safe_div(total_cost, device_count),
                        safe_div(var_cost, device_count),
                        safe_div(fix_cost, device_count),
                        safe_div(sums['영업이익'], device_count)
                    ]
                    rows.append(vals)

        monthly_tables.append(rows)

    if not monthly_tables:
        logger.warning("⚠️ 유효한 월별 원가 데이터가 없습니다.")
        return []

    logger.info(f"월별 데이터 병합 중: {len(monthly_tables)}개월 데이터")
    arr = np.array(monthly_tables)
    avg_table = np.nanmean(arr, axis=0)

    out_table = []
    for idx in range(5):  # 5개 컬럼만
        row_data = [avg_table[i][idx] for i in range(len(avg_table))]
        out_table.append(row_data)

    logger.info(f"✅ 6번 표 (원가현황) 생성 완료: {len(out_table)}행 x 5열")
    return out_table
def make_average_financial_data_per_device(df, logger):
    """6번 표: 채널별 원가현황 (월별 평균 계산, 점주대행 제외)
    - 출력: 매출, 총비용, 변동비, 고정비, 영업이익, 이익율 (6개 row)
    - 편의점, 브랜드, 일반장소 × ATM/CD 조합 (총 5열)
    """
    logger.info("💰 6번 표 생성: 채널별 원가현황 (월별 평균 계산, 점주대행 제외)")
    
    groups = ['편의점', '브랜드', '일반장소']
    devices = ['ATM', 'CD']
    months = sorted(df['month'].dropna().unique())
    
    logger.info(f"분석 대상 월: {len(months)}개월")

    monthly_tables = []

    for m in months:
        cur_month = pd.to_datetime(m)
        logger.info(f"월별 데이터 처리 중: {cur_month.strftime('%Y-%m')}")

        df_active = df[(df['month'] == cur_month) & (df['최초설치일'] <= cur_month)].copy()
        if df_active.empty:
            logger.warning(f"{cur_month.strftime('%Y-%m')}: 활성 데이터 없음")
            continue

        df_active['그룹'] = df_active.apply(lambda row: (
            '편의점' if row['중분류1'] in MAJOR_BRANDS else
            '브랜드' if row['대분류1'] in ['브랜드', '브랜드제휴'] else
            '일반장소' if row['대분류1'] in ['일반장소', '일반(기타기관+개별장소)'] else
            None
        ), axis=1)
        df_active = df_active[df_active['그룹'].notna() & df_active['CD/ATM'].isin(['ATM', 'CD'])]

        logger.info(f"{cur_month.strftime('%Y-%m')}: 유효 데이터 {len(df_active)}행")

        rows = []
        for group in groups:
            if group == '브랜드':
                sub_atm = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == 'ATM')]
                sub_cd  = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == 'CD')]
                sub = pd.concat([sub_atm, sub_cd])
                device_count = sub['기번1'].nunique()
                if device_count == 0:
                    rows.append([np.nan] * 6)
                    continue

                sums = sub[['매출합','영업이익','임차료','현송비','중계/명세','유지보수','회선비',
                            '장애대처/경비','수선공사/청소','감가상각비','기타고정비','인건비/경비','준고정비']].sum(numeric_only=True)

                var_cost = sums[['임차료','현송비','중계/명세','유지보수','회선비','장애대처/경비','수선공사/청소']].sum()
                fix_cost = sums[['감가상각비','기타고정비','인건비/경비','준고정비']].sum()
                total_cost = var_cost + fix_cost
                sale = sums['매출합']
                op_profit = sums['영업이익']

                vals = [
                    safe_div(sale, device_count),
                    safe_div(total_cost, device_count),
                    safe_div(var_cost, device_count),
                    safe_div(fix_cost, device_count),
                    safe_div(op_profit, device_count),
                    round(safe_div(op_profit, sale) * 100)  # 이익율 (%)
                ]
                rows.append(vals)
            else:
                for device in devices:
                    sub = df_active[(df_active['그룹'] == group) & (df_active['CD/ATM'] == device)]
                    device_count = sub['기번1'].nunique()
                    if device_count == 0:
                        rows.append([np.nan] * 6)
                        continue

                    sums = sub[['매출합','영업이익','임차료','현송비','중계/명세','유지보수','회선비',
                                '장애대처/경비','수선공사/청소','감가상각비','기타고정비','인건비/경비','준고정비']].sum(numeric_only=True)

                    var_cost = sums[['임차료','현송비','중계/명세','유지보수','회선비','장애대처/경비','수선공사/청소']].sum()
                    fix_cost = sums[['감가상각비','기타고정비','인건비/경비','준고정비']].sum()
                    total_cost = var_cost + fix_cost
                    sale = sums['매출합']
                    op_profit = sums['영업이익']

                    vals = [
                        safe_div(sale, device_count),
                        safe_div(total_cost, device_count),
                        safe_div(var_cost, device_count),
                        safe_div(fix_cost, device_count),
                        safe_div(op_profit, device_count),
                        round(safe_div(op_profit, sale) * 100)  # 이익율 (%)
                    ]
                    rows.append(vals)

        monthly_tables.append(rows)

    if not monthly_tables:
        logger.warning("⚠️ 유효한 월별 원가 데이터가 없습니다.")
        return []

    logger.info(f"월별 데이터 병합 중: {len(monthly_tables)}개월 데이터")
    arr = np.array(monthly_tables)
    avg_table = np.nanmean(arr, axis=0)

    out_table = []
    for idx in range(6):  # ✅ 이익율 포함해서 6개 row
        row_data = [avg_table[i][idx] for i in range(len(avg_table))]
        out_table.append(row_data)

    logger.info(f"✅ 6번 표 (원가현황) 생성 완료: {len(out_table)}행 x {len(out_table[0])}열")
    return out_table


def make_financial_data(df, logger):
    """6번째 표: 채널별 원가현황 - 대수 + 건당수수료 계산
    - 편의점, 일반장소는 ATM, CD로 나눠 계산
    - 브랜드와 브랜드제휴는 합산하여 ATM, CD 구분 없이 계산
    - 점주대행은 출력 제외 (H열은 건드리지 않음)
    """
    logger.info("💰 6번 표 생성: 대수 및 건당수수료 계산 (점주대행 컬럼 제외)")
    
    def map_group(row):
        if row.get('중분류1') in MAJOR_BRANDS:
            return '편의점'
        dv = row.get('대분류1')
        if dv in ['브랜드', '브랜드제휴']:
            return '브랜드'
        if dv in ['일반장소', '일반(기타기관+개별장소)']:
            return '일반장소'
        return None  # 점주대행 제외

    tmp = df.copy()
    tmp['그룹'] = tmp.apply(map_group, axis=1)
    tmp = tmp[tmp['그룹'].notna() & tmp['CD/ATM'].isin(['ATM', 'CD'])]
    
    logger.info(f"필터링 후 데이터: {len(tmp)}행")

    # 🔑 5개 컬럼만 출력 (점주대행 제외)
    group_device_keys = [
        ('편의점', 'ATM'),
        ('편의점', 'CD'),
        ('브랜드',  '합산'),   # ✅ 브랜드는 ATM+CD 합산
        ('일반장소', 'ATM'),
        ('일반장소', 'CD')
    ]

    data_rows = []

    # --- (1) 대수 계산 ---
    logger.info("📊 그룹별 대수 계산 중...")
    dev_counts = tmp.groupby(['그룹', 'CD/ATM'])['기번1'].nunique()
    row1 = []
    for group, device in group_device_keys:
        if device == '합산':
            val = dev_counts.get((group, 'ATM'), 0) + dev_counts.get((group, 'CD'), 0)
            logger.info(f"{group} 합산: {val}대 (ATM: {dev_counts.get((group, 'ATM'), 0)}, CD: {dev_counts.get((group, 'CD'), 0)})")
        else:
            val = dev_counts.get((group, device), 0)
            logger.info(f"{group} {device}: {val}대")
        row1.append(val)
    data_rows.append(row1)

    # --- (2) 건당 수수료 계산 ---
    logger.info("💰 그룹별 건당수수료 계산 중...")
    per_dev = tmp.groupby(['그룹', 'CD/ATM', '기번1'])
    fee_sum = per_dev[['수수료매출']].sum().groupby(level=[0, 1]).sum()
    cnt_sum = per_dev['일건수'].mean().groupby(level=[0, 1]).sum()
    df_combined = pd.concat([fee_sum, cnt_sum.rename('일건수')], axis=1)

    row2 = []
    for group, device in group_device_keys:
        if device == '합산':
            s1 = df_combined.loc[(group, 'ATM')] if (group, 'ATM') in df_combined.index else pd.Series({'수수료매출': 0, '일건수': 0})
            s2 = df_combined.loc[(group, 'CD')] if (group, 'CD') in df_combined.index else pd.Series({'수수료매출': 0, '일건수': 0})
            total_fee = s1['수수료매출'] + s2['수수료매출']
            total_cnt = s1['일건수'] + s2['일건수']
            val = safe_div(total_fee, total_cnt * 365)
            logger.info(f"{group} 합산 건당수수료: {val:.2f}원" if not pd.isna(val) else f"{group} 합산 건당수수료: N/A")
        else:
            if (group, device) in df_combined.index:
                s = df_combined.loc[(group, device)]
                val = safe_div(s['수수료매출'], s['일건수'] * 365)
                logger.info(f"{group} {device} 건당수수료: {val:.2f}원" if not pd.isna(val) else f"{group} {device} 건당수수료: N/A")
            else:
                val = None
                logger.warning(f"{group} {device}: 데이터 없음")
        row2.append(val if not pd.isna(val) else "")
    data_rows.append(row2)

    logger.info(f"✅ 6번 표 (대수/건당수수료) 생성 완료: {len(data_rows)}행 x 5열")
    return data_rows


# ===== main 함수 =====
def main():
    """
    - 1번 표(운영현황)에는 점주대행 카운팅 포함
    - 6번 표(채널별 원가현황)에는 점주대행 완전히 제외
    - 표별 group mapping 함수 따로 사용
    """
    logger = setup_logging()
    
    logger.info("🚀 메인 프로세스 시작")
    logger.info(f"파일 경로: {FILE_PATH}")
    logger.info(f"입력 시트: {SHEET_IN}, 출력 시트: {SHEET_OUT}")

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        logger.info("📂 Excel 파일 열기...")
        wb = xw.Book(FILE_PATH)
        ws_in = wb.sheets[SHEET_IN]
        ws_out = wb.sheets[SHEET_OUT] if SHEET_OUT in [s.name for s in wb.sheets] else wb.sheets.add(SHEET_OUT)
        logger.info("✅ Excel 파일 열기 완료")

        # ----- 데이터 로드 및 전처리 -----
        df = load_and_clean_data(ws_in, logger)
        
        # 날짜 전처리
        logger.info("📅 날짜 컬럼 전처리 시작")
        if 'month' in df.columns:
            logger.info("월 컬럼 변환 중...")
            df['month'] = pd.to_datetime(
                '2024-' + df['month'].astype(str).str.replace('월', '').str.zfill(2),
                format='%Y-%m',
                errors='coerce'
            )
            valid_months = df['month'].notna().sum()
            logger.info(f"유효한 월 데이터: {valid_months}개")

        if '최초설치일' in df.columns:
            logger.info("최초설치일 변환 중...")
            df['최초설치일'] = pd.to_datetime(df['최초설치일'], errors='coerce')
            valid_install_dates = df['최초설치일'].notna().sum()
            logger.info(f"유효한 설치일 데이터: {valid_install_dates}개")
            if valid_install_dates > 0:
                logger.info(f"설치일 범위: {df['최초설치일'].min()} ~ {df['최초설치일'].max()}")

        logger.info("✅ 날짜 전처리 완료")

        # ----- 1번 표(운영현황)용: 점주대행 포함 -----
        logger.info("🔄 1번 표용 데이터 준비 (점주대행 포함)")
        df1 = df.copy()
        df1['그룹'] = df1.apply(map_group_for_summary, axis=1)
        df1 = df1[df1['그룹'].notna()]
        group_counts = df1['그룹'].value_counts()
        logger.info(f"1번 표 그룹별 데이터 수: {group_counts.to_dict()}")

        # ----- 6번 표(원가현황)용: 점주대행 제외 -----
        logger.info("🔄 6번 표용 데이터 준비 (점주대행 제외)")
        df6 = df.copy()
        df6['그룹'] = df6.apply(map_group_for_financial, axis=1)
        df6 = df6[df6['그룹'].notna()]
        group_counts_6 = df6['그룹'].value_counts()
        logger.info(f"6번 표 그룹별 데이터 수: {group_counts_6.to_dict()}")

        # ----- 각 표 생성 -----
        logger.info("📋 각 표 생성 시작...")
        
        data1 = make_summary_data(df1, logger)
        data2 = make_convenience_data(df, logger)
        data3 = make_grade_data(df, logger)
        data4 = make_location_maker_data(df, logger)
        data5 = make_issue_data(df, logger)
        data6_part1 = make_financial_data(df6, logger)               # [대수, 건당수수료]
        data6_part2 = make_average_financial_data_per_device(df6, logger)  # [매출 ~ 이익율]

        logger.info("📊 모든 표 생성 완료")

        # ----- 6번 표 합치기 (점주대행 컬럼 제외, 5개 컬럼만) -----
        logger.info("🔧 6번 표 데이터 병합 및 포맷팅...")
        
        def ensure_n_cols(row, n=5):
            return row[:n] if len(row) > n else row + [""] * (n - len(row))

        # 정확히 C61:G68 (8행 × 5열)에 맞춤
        data6 = [ensure_n_cols(data6_part1[0])] + \
                [ensure_n_cols(r) for r in data6_part2] + \
                [[""] * 5] + \
                [ensure_n_cols(data6_part1[1])]
        
        # [대수] + [5개 row(매출~영업이익)] + [건당수수료]
        data6 = [data6_part1[0]] + data6_part2 + [data6_part1[1]]
        
        logger.info(f"6번 표 최종 구성: {len(data6)}행 x 5열")

        # ----- 엑셀에 각 표 반영 -----
        logger.info("📝 Excel 시트에 데이터 쓰기 시작...")
        
        datasets = {
            'final': data1,
            'final2': data2,
            'final3': data3,
            'final4': data4,
            'final5': data5,
            'final6': data6
        }

        for key, data in datasets.items():
            if not data:
                logger.warning(f"⚠️ {key} 테이블 데이터 없음 - 건너뛰기")
                continue
                
            logger.info(f"📊 {key} 테이블 처리 중...")
            
            # 데이터 무결성 검사
            row_lengths = [len(row) for row in data if hasattr(row, '__len__')]
            if len(set(row_lengths)) != 1:
                logger.error(f"❌ {key} 테이블 행 길이 불일치: {row_lengths}")
                continue
                
            start_cell = DATA_START_CELLS[key]
            nrows = len(data)
            ncols = len(data[0]) if data else 0
            
            logger.info(f"{key}: {nrows}행 x {ncols}열, 시작셀: {start_cell}")
            
            if nrows > 0 and ncols > 0:
                # 6번 표는 특별히 G열까지만 (C61:G68)
                if key == 'final6':
                    data_range = ws_out.range(f"{start_cell}:{FINAL6_END_CELL}")
                    logger.info(f"{key}: 특별 범위 설정 ({start_cell}:{FINAL6_END_CELL}) - 점주대행 컬럼(H) 보호")
                else:
                    end_cell = ws_out.range(start_cell).offset(nrows - 1, ncols - 1).address
                    data_range = ws_out.range(f"{start_cell}:{end_cell}")
                
                try:
                    data_range.clear_contents()
                    logger.info(f"{key}: 기존 내용 삭제 완료")
                except:
                    try:
                        data_range.api.UnMerge()
                        data_range.clear_contents()
                        logger.info(f"{key}: 병합 해제 후 내용 삭제 완료")
                    except Exception as e:
                        logger.warning(f"{key}: 내용 삭제 실패 - {e}")
                
                try:
                    ws_out.range(start_cell).value = data
                    logger.info(f"✅ {key} 테이블 쓰기 완료")
                except Exception as e:
                    logger.error(f"❌ {key} 테이블 쓰기 실패: {e}")

        logger.info("💾 Excel 파일 저장 중...")
        wb.save()
        
        # [대수] + [6개 row(매출~이익율)] + [건당수수료]
        data6 = [data6_part1[0]] + data6_part2 + [data6_part1[1]]
        
        # 무조건 앞 6개 열만 유지
        data6 = [r[:6] for r in data6]
        logger.info(f"6번 표 최종 구성: {len(data6)}행 x 6열")

        # ----- 엑셀에 각 표 반영 -----
        logger.info("📝 Excel 시트에 데이터 쓰기 시작...")
        
        datasets = {
            'final': data1,
            'final2': data2,
            'final3': data3,
            'final4': data4,
            'final5': data5,
            'final6': data6
        }

        for key, data in datasets.items():
            if not data:
                logger.warning(f"⚠️ {key} 테이블 데이터 없음 - 건너뛰기")
                continue
                
            logger.info(f"📊 {key} 테이블 처리 중...")
            
            # 데이터 무결성 검사
            row_lengths = [len(row) for row in data if hasattr(row, '__len__')]
            if len(set(row_lengths)) != 1:
                logger.error(f"❌ {key} 테이블 행 길이 불일치: {row_lengths}")
                continue
                
            start_cell = DATA_START_CELLS[key]
            nrows = len(data)
            ncols = len(data[0]) if data else 0
            
            logger.info(f"{key}: {nrows}행 x {ncols}열, 시작셀: {start_cell}")
            
            if nrows > 0 and ncols > 0:
                end_cell = ws_out.range(start_cell).offset(nrows - 1, ncols - 1).address
                data_range = ws_out.range(f"{start_cell}:{end_cell}")
                
                try:
                    data_range.clear_contents()
                    logger.info(f"{key}: 기존 내용 삭제 완료")
                except:
                    try:
                        data_range.api.UnMerge()
                        data_range.clear_contents()
                        logger.info(f"{key}: 병합 해제 후 내용 삭제 완료")
                    except Exception as e:
                        logger.warning(f"{key}: 내용 삭제 실패 - {e}")
                
                try:
                    ws_out.range(start_cell).value = data
                    logger.info(f"✅ {key} 테이블 쓰기 완료")
                except Exception as e:
                    logger.error(f"❌ {key} 테이블 쓰기 실패: {e}")

        logger.info("💾 Excel 파일 저장 중...")
        wb.save()
        logger.info("✅ Excel 파일 저장 완료")
        
    except Exception as e:
        logger.error(f"❌ 처리 중 오류 발생: {e}")
        import traceback
        logger.error(f"상세 오류:\n{traceback.format_exc()}")
        raise
        
    finally:
        try:
            wb.close()
            logger.info("📂 Excel 파일 닫기 완료")
        except:
            logger.warning("⚠️ Excel 파일 닫기 실패")
        
        app.quit()
        logger.info("🔚 Excel 애플리케이션 종료")

    logger.info("🎉 === ATM 데이터 처리 완료 ===")


if __name__ == "__main__":
    main()
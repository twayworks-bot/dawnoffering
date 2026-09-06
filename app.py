import os
import re
import io
import urllib.parse
from datetime import datetime, date, timedelta
from flask import Flask, Blueprint, render_template, request, redirect, url_for
import pandas as pd
import requests

app = Flask(__name__)

# 블루프린트를 이용한 Prefix (/dawn_offering) 구현
bp = Blueprint('dawn_offering', __name__, url_prefix='/dawn_offering')

# .env 파일 로드 함수
def load_env():
    env = {}
    if os.path.exists('.env'):
        with open('.env', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    if '=' in line:
                        k, v = line.split('=', 1)
                        # 따옴표 제거
                        val = v.strip().strip('"').strip("'")
                        env[k.strip()] = val
    return env

env_config = load_env()
# 환경 변수 우선 및 .env 파일 로드 폴백 설정 (도커 컨테이너화 지원)
DEFAULT_URL = os.environ.get('GOGOLE_XLSX_URL', env_config.get('GOGOLE_XLSX_URL', ''))
DEFAULT_SHEET = os.environ.get('XLSX_SHEETNAME', env_config.get('XLSX_SHEETNAME', '순서표원천'))
DEFAULT_MUSIC_SHEET = os.environ.get('WORSHIP_MUSIC_SHEETNAME', env_config.get('WORSHIP_MUSIC_SHEETNAME', '영상일정표'))

# 자동 열 인식을 위한 키워드 정의
KEYWORDS = {
    'date': ['날짜', '일자', 'date', '일기'],
    'leader': ['인도자', '인도', 'leader'],
    'prayer1': ['기도1', '기도자1', '회개', 'prayer1'],
    'prayer2': ['기도2', '기도자2', '교회', 'prayer2'],
    'broadcast': ['방송지원', '방송', '음향', 'media', 'broadcast'],
    'guide': ['안내자', '안내', 'usher', 'guide'],
    'praise': ['찬양', '찬송', '찬송가', 'praise', 'hymn']
}

def convert_google_drive_url(url):
    """
    구글 드라이브/시트 공유 링크를 직접 다운로드 가능한 Excel(xlsx) Export URL로 변환합니다.
    시트 이름 지정을 지원하기 위해 엑셀 형식으로 내보내기를 수행합니다.
    """
    url = url.strip()
    
    # 1. 구글 스프레드시트 주소인 경우
    sheet_match = re.search(r'docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    if sheet_match:
        spreadsheet_id = sheet_match.group(1)
        # XLSX 형식으로 내보내기 수행 (전체 시트를 받기 위함)
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx", "excel"
            
    # 2. 구글 드라이브 업로드된 파일 (엑셀) 주소인 경우
    file_match = re.search(r'drive\.google\.com/file/d/([a-zA-Z0-9-_]+)', url)
    if file_match:
        file_id = file_match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}", "excel"
        
    # 3. 기타 직링크 처리
    if url.lower().endswith('.csv'):
        return url, "csv"
    if url.lower().endswith(('.xlsx', '.xls')):
        return url, "excel"
        
    # 기본값은 구글 시트 엑셀 내보내기 형식으로 취급
    return url, "excel"

def fetch_data(url, url_type, sheet_name=None):
    """
    변환된 URL을 바탕으로 서버에서 데이터를 안전하게 가져와 DataFrame으로 로드합니다.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    
    if url_type == "csv":
        try:
            content = response.content.decode('utf-8')
            return pd.read_csv(io.StringIO(content))
        except UnicodeDecodeError:
            content = response.content.decode('cp949', errors='ignore')
            return pd.read_csv(io.StringIO(content))
    else:
        # Excel: 시트 이름이 지정되어 있으면 해당 시트를 파싱하고, 실패하면 첫 번째 시트 로드
        xl = pd.ExcelFile(io.BytesIO(response.content))
        if sheet_name and sheet_name in xl.sheet_names:
            return xl.parse(sheet_name)
        elif len(xl.sheet_names) > 0:
            return xl.parse(0)
        else:
            raise ValueError("엑셀 파일에 시트가 존재하지 않습니다.")

def find_column(df, keywords):
    """
    주어진 키워드들을 포함하는 데이터프레임 컬럼 이름을 반환합니다.
    """
    for col in df.columns:
        col_str = str(col).strip().lower()
        for kw in keywords:
            if kw.lower() in col_str:
                return col
    return None

def map_dataframe_to_columns(df):
    """
    데이터프레임의 실제 열 이름을 사전 약속된 기능별로 매핑합니다. (동적 매핑 + 고정 위치 폴백)
    """
    mapping = {}
    
    # 1. 키워드 기반 동적 열 매칭 시도
    for key, keywords in KEYWORDS.items():
        mapping[key] = find_column(df, keywords)
        
    # 2. 매칭에 실패한 경우, 순서에 따른 폴백 매핑 적용 (A~G열)
    position_fallback = {
        'date': 0,
        'leader': 1,
        'prayer1': 2,
        'prayer2': 3,
        'broadcast': 4,
        'guide': 5,
        'praise': 6
    }
    
    num_cols = len(df.columns)
    for key, fallback_idx in position_fallback.items():
        if mapping[key] is None:
            if fallback_idx < num_cols:
                mapping[key] = df.columns[fallback_idx]
                
    return mapping

def get_target_week_dates(ref_date):
    """
    기준 날짜(datetime.date)를 토대로 출력할 월~금요일 날짜 리스트를 도출합니다.
    - 월~금요일: 당주 월~금요일
    - 토~일요일: 차주 월~금요일
    """
    wd = ref_date.weekday() # 0=월, 1=화, ... 5=토, 6=일
    
    if wd <= 4:
        # 평일: 당주 월요일 기점
        monday = ref_date - timedelta(days=wd)
    else:
        # 주말: 다음 주 월요일 기점
        days_until_next_mon = 7 - wd
        monday = ref_date + timedelta(days=days_until_next_mon)
        
    return [monday + timedelta(days=i) for i in range(5)]

def find_week_header_row(df, d_mon):
    """
    주어진 월요일 날짜(d_mon)에 매칭되는 주차 구분 헤더 행의 인덱스를 찾습니다.
    """
    col_0 = df.columns[0]
    
    # 다양한 날짜 표기 형식 생성
    patterns = [
        f"{d_mon.year:04d}-{d_mon.month:02d}-{d_mon.day:02d}", # 2026-09-07
        f"{d_mon.year:04d}.{d_mon.month:02d}.{d_mon.day:02d}", # 2026.09.07
        f"{d_mon.year:04d}/{d_mon.month:02d}/{d_mon.day:02d}", # 2026/09/07
        f"{d_mon.month}/{d_mon.day}",                           # 9/7
        f"{d_mon.month:02d}/{d_mon.day:02d}",                     # 09/07
        f"{d_mon.month}.{d_mon.day}",                           # 9.7
        f"{d_mon.month:02d}.{d_mon.day:02d}",                     # 09.07
        f"{d_mon.month}월 {d_mon.day}일",                         # 9월 7일
        f"{d_mon.month}월{d_mon.day}일"                          # 9월7일
    ]
    
    for idx, val in enumerate(df[col_0]):
        if pd.isna(val):
            continue
            
        # 1. 셀 자체가 datetime 또는 Timestamp 타입인 경우
        if isinstance(val, (datetime, pd.Timestamp)):
            if val.date() == d_mon:
                return idx
                
        # 2. 문자열 비교 및 부분 검색
        val_str = str(val).strip()
        if any(p in val_str for p in patterns):
            return idx
            
        # 3. 문자열 날짜를 직접 파싱하여 비교 시도
        try:
            parsed_date = pd.to_datetime(val_str).date()
            if parsed_date == d_mon:
                return idx
        except Exception:
            pass
            
    return None

def normalize_date(val, ref_year):
    """
    데이터프레임 셀 값의 다양한 포맷을 datetime.date 객체로 변환합니다.
    """
    if pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()
    
    val_str = str(val).strip()
    val_str = re.sub(r'[월화수목금토일]요일', '', val_str) # 요일 텍스트 제거
    
    # 1. 숫자만 추출하여 YYYY-MM-DD 또는 MM-DD 분석
    digits = re.findall(r'\d+', val_str)
    if len(digits) >= 3:
        try:
            year = int(digits[0])
            if year < 100:
                year += 2000
            month = int(digits[1])
            day = int(digits[2])
            return datetime(year, month, day).date()
        except ValueError:
            pass
    elif len(digits) == 2:
        try:
            month = int(digits[0])
            day = int(digits[1])
            return datetime(ref_year, month, day).date()
        except ValueError:
            pass
            
    # 2. 판다스 기본 파서 사용 시도
    try:
        return pd.to_datetime(val_str).date()
    except Exception:
        pass
        
    return None

def find_worship_music_folder_url(df_music, d_mon):
    """
    '영상일정표' 시트에서 기준 월요일 날짜(d_mon)에 해당되는 구글 드라이브 폴더 링크를 추출합니다.
    """
    patterns = [
        f"{d_mon.month}/{d_mon.day}-주차",
        f"{d_mon.month:02d}/{d_mon.day:02d}-주차",
        f"{d_mon.month}.{d_mon.day}-주차",
        f"{d_mon.month:02d}.{d_mon.day:02d}-주차",
        f"{d_mon.month}월 {d_mon.day}일-주차",
        f"{d_mon.month}월{d_mon.day}일-주차"
    ]
    
    for col in df_music.columns:
        for idx, val in enumerate(df_music[col]):
            if pd.isna(val):
                continue
            val_str = str(val).strip()
            if any(p in val_str for p in patterns):
                # 찾은 인덱스의 바로 다음 행에 있는 구글 드라이브 폴더 주소를 로드
                if idx + 1 < len(df_music):
                    next_val = df_music[col].iloc[idx + 1]
                    if pd.notna(next_val):
                        next_val_str = str(next_val).strip()
                        if next_val_str.startswith("http"):
                            return next_val_str
    return None

def scrape_folder_files(folder_url):
    """
    구글 드라이브 공유 폴더 페이지를 크롤링하여 포함된 미디어 파일들의 목록을 읽어옵니다.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    }
    try:
        r = requests.get(folder_url, headers=headers, timeout=15)
        if r.status_code != 200:
            return []
        
        # 따옴표로 감싸진 [1-5].요일 로 시작하는 오디오/비디오 파일명 추출
        files = re.findall(r'"([1-5]\.[^"]+\.(?:mp4|mp3|m4a|wav|avi|mkv))"', r.text)
        # JSON slash 이스케이프 제거
        files = [f.replace(r'\/', '/') for f in files]
        return sorted(list(set(files)))
    except Exception:
        return []

def extract_chanyang_title(filename):
    """
    파일 명에서 요일 접두사 및 .YYYYMMDD. 접미사를 트리밍하여 순수 찬양 제목을 추출합니다.
    """
    # 1. 날짜로 시작하는 .YYYYMMDD. 뒤에 붙은 확장자 및 주일폐회찬양 등 접미사 제거
    match_suffix = re.search(r'\.\d{8}\.', filename)
    if match_suffix:
        idx = match_suffix.start()
        title_part = filename[:idx]
    else:
        title_part = re.sub(r'\.(?:mp4|mp3|m4a|wav|avi|mkv)$', '', filename, flags=re.IGNORECASE)
        title_part = title_part.replace('.주일폐회찬양', '').replace('.찬양', '')
        
    # 2. 접두사 "번호.요일." 트리밍 (예: 1.월., 2.화.)
    title_part = re.sub(r'^[1-5]\.[월화수목금]\.', '', title_part)
    
    # 3. 추가 일련 번호 트리밍 (예: 157., 080.)
    title_part = re.sub(r'^\d+\.', '', title_part)
    
    return title_part.strip(' .')

@bp.route('/')
def index():
    error = request.args.get('error')
    # 기본 날짜를 2026-09-06으로 설정
    default_date = "2026-09-06"
    return render_template('index.html', error=error, default_url=DEFAULT_URL, default_date=default_date)

@bp.route('/api/status')
def status():
    """
    도커 컨테이너 헬스체크용 엔드포인트
    """
    return {"status": "healthy", "sheet_name": DEFAULT_SHEET, "music_sheet_name": DEFAULT_MUSIC_SHEET}, 200

@bp.route('/generate', methods=['POST'])
def generate():
    url = request.form.get('url')
    date_str = request.form.get('date')
    
    if not url or not date_str:
        return redirect(url_for('dawn_offering.index', error="모든 값을 입력해주세요."))
        
    try:
        ref_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return redirect(url_for('dawn_offering.index', error="날짜 포맷이 올바르지 않습니다. (YYYY-MM-DD 형식을 사용해야 합니다)"))
        
    # 출력 대상이 될 5개의 날짜 구하기 (월~금)
    target_dates = get_target_week_dates(ref_date)
    d_mon = target_dates[0]
    
    # URL 분석 및 파일 로드 (.env에 선언된 XLSX_SHEETNAME 전달)
    try:
        direct_url, url_type = convert_google_drive_url(url)
        # 두 개의 시트(순서표원천, 영상일정표)를 모두 안전하게 추출하기 위해 전체 Excel 객체를 직접 컨트롤
        xl = pd.ExcelFile(io.BytesIO(requests.get(direct_url, timeout=15).content))
        
        # 1. 예배 순서 시트 파싱
        if DEFAULT_SHEET in xl.sheet_names:
            df = xl.parse(DEFAULT_SHEET)
        else:
            df = xl.parse(0)
            
        # 2. 찬양 음악 시트 파싱 및 드라이브 파일 스크래핑 연동
        chanyang_map = {}
        if DEFAULT_MUSIC_SHEET in xl.sheet_names:
            try:
                df_music = xl.parse(DEFAULT_MUSIC_SHEET)
                # 구글 드라이브 폴더 URL 검색
                folder_url = find_worship_music_folder_url(df_music, d_mon)
                if folder_url:
                    # 폴더 내에 등록된 라이브 파일 리스트 로드
                    drive_files = scrape_folder_files(folder_url)
                    
                    # 요일별 접두사 매핑 정의
                    weekday_prefixes = {
                        0: "1.월.",
                        1: "2.화.",
                        2: "3.수.",
                        3: "4.목.",
                        4: "5.금."
                    }
                    
                    for idx, d in enumerate(target_dates):
                        prefix = weekday_prefixes[idx]
                        # 접두사 매칭 + 찬양 키워드 포함 파일 필터링
                        matched_chanyangs = [
                            f for f in drive_files 
                            if f.startswith(prefix) and ("찬양" in f or "주일폐회찬양" in f)
                        ]
                        
                        if matched_chanyangs:
                            chanyang_map[d] = extract_chanyang_title(matched_chanyangs[0])
                        else:
                            chanyang_map[d] = "미업로드"
            except Exception as music_err:
                print(f"음악 데이터 로딩/크롤링 실패: {music_err}")
                
    except Exception as e:
        return redirect(url_for('dawn_offering.index', error=f"구글 링크에서 데이터를 불러오지 못했습니다. 링크 공유 설정 및 URL 주소를 다시 한 번 확인해 주세요. 상세 내용: {str(e)}"))
        
    if df.empty:
        return redirect(url_for('dawn_offering.index', error="불러온 엑셀/시트의 데이터가 비어 있습니다."))
        
    schedule = []
    
    # [우선순위 전략 A] "순서표원천"에 최적화된 주차 헤더(예: 09/07 ~ 09/11) 매칭 및 오프셋 파싱
    header_idx = find_week_header_row(df, d_mon)
            
    # 주차 헤더를 찾았고 그 아래 5일치 이상의 행 데이터가 확보된 경우
    if header_idx is not None and header_idx + 5 < len(df):
        for i, d in enumerate(target_dates):
            row = df.iloc[header_idx + 1 + i]
            disp_date_str = f"{d.strftime('%Y. %m. %d.')} {'월화수목금토일'[d.weekday()]}요일"
            tab_label = f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})"
            
            # B열(1)=인도자, C열(2)=기도자1, D열(3)=기도자2, E열(4)=방송지원, F열(5)=안내자
            leader = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            prayer1 = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
            prayer2 = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
            broadcast = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
            guide = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ""
            
            # 찬양명 매핑 (구글드라이브 폴더 우선 연동, 실패 시 엑셀 원천 데이터 폴백)
            if d in chanyang_map:
                praise = chanyang_map[d]
            else:
                praise = str(row.iloc[6]).strip() if len(row) > 6 and pd.notna(row.iloc[6]) else "예수 하나님의 공의 - 예수 하나님의 공의"
                if not praise or praise == "불일치":
                    praise = "미업로드"
                    
            schedule.append({
                'date': d,
                'date_str': disp_date_str,
                'tab_label': tab_label,
                'leader': leader,
                'guide': guide,
                'broadcast': broadcast,
                'prayer1': prayer1,
                'prayer2': prayer2,
                'praise': praise
            })
    else:
        # [폴백 전략 B] 기존의 개별 행 날짜 직접 일치 방식
        mapped_cols = map_dataframe_to_columns(df)
        
        if mapped_cols['date'] is None:
            return redirect(url_for('dawn_offering.index', error="시트 내부에서 '날짜' 또는 '일자' 열을 인식할 수 없습니다."))
            
        for d in target_dates:
            match_row = None
            for _, row in df.iterrows():
                row_date = normalize_date(row[mapped_cols['date']], ref_date.year)
                if row_date == d:
                    match_row = row
                    break
                    
            disp_date_str = f"{d.strftime('%Y. %m. %d.')} {'월화수목금토일'[d.weekday()]}요일"
            tab_label = f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})"
            
            if match_row is not None:
                leader = str(match_row[mapped_cols['leader']]).strip() if pd.notna(match_row[mapped_cols['leader']]) else ""
                guide = str(match_row[mapped_cols['guide']]).strip() if pd.notna(match_row[mapped_cols['guide']]) else ""
                broadcast = str(match_row[mapped_cols['broadcast']]).strip() if pd.notna(match_row[mapped_cols['broadcast']]) else ""
                prayer1 = str(match_row[mapped_cols['prayer1']]).strip() if pd.notna(match_row[mapped_cols['prayer1']]) else ""
                prayer2 = str(match_row[mapped_cols['prayer2']]).strip() if pd.notna(match_row[mapped_cols['prayer2']]) else ""
                
                if d in chanyang_map:
                    praise = chanyang_map[d]
                else:
                    praise = str(match_row[mapped_cols['praise']]).strip() if mapped_cols['praise'] in match_row and pd.notna(match_row[mapped_cols['praise']]) else "예수 하나님의 공의 - 예수 하나님의 공의"
                    if not praise or praise == "불일치":
                        praise = "미업로드"
            else:
                leader = ""
                guide = ""
                broadcast = ""
                prayer1 = ""
                prayer2 = ""
                praise = chanyang_map.get(d, "미업로드")
                
            schedule.append({
                'date': d,
                'date_str': disp_date_str,
                'tab_label': tab_label,
                'leader': leader,
                'guide': guide,
                'broadcast': broadcast,
                'prayer1': prayer1,
                'prayer2': prayer2,
                'praise': praise
            })
        
    return render_template('dawn_prayers.html', schedule=schedule)

# 블루프린트 등록
app.register_blueprint(bp)

# 루트 주소(/)로 오면 Prefix(/dawn_offering/)로 리다이렉트
@app.route('/')
def root():
    return redirect('/dawn_offering/')

if __name__ == '__main__':
    # 5000번 포트로 구동
    app.run(host='0.0.0.0', port=5000, debug=True)

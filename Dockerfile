# 1. Base Image - 경량화된 공식 Python 이미지 사용
FROM python:3.13-slim

# 2. 환경변수 설정 - Python 최적화 및 버퍼 해제
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=5000

# 3. .env 설정을 기본 환경변수로 컨테이너 내부 정의 (호스트 기동 시 -e 옵션으로 동적 주입 가능)
ENV GOGOLE_XLSX_URL="https://docs.google.com/spreadsheets/d/1RWKAewSqTB6wVCw_-R86NjoGszd0Y1Sy/edit?usp=drive_link&ouid=108300514251400381497&rtpof=true&sd=true" \
    XLSX_SHEETNAME="순서표원천" \
    WORSHIP_MUSIC_SHEETNAME="영상일정표"

# 4. 컨테이너 내부 작업 디렉토리 지정
WORKDIR /app

# 5. 종속성 설치 패키지 복사 및 다운로드
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 6. 애플리케이션 전체 소스 복사
COPY . .

# 7. 외부 노출 포트 지정
EXPOSE 5000

# 8. 컨테이너 헬스체크 정의 (요청한 사양에 맞춘 Prefix 경로 연동)
HEALTHCHECK --interval=30s --timeout=10s --retries=3 --start-period=5s \
    CMD python -c "import requests; requests.get('http://localhost:5000/dawn_offering/api/status')"

# 9. Gunicorn 운영서버로 기동 (요청한 사양에 맞춘 최적화 워커 및 디버깅 로그 옵션)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--log-level", "debug", "--access-logfile", "-", "--error-logfile", "-", "--capture-output", "app:app"]

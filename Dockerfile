FROM debian:trixie-slim

# 파이썬 + LibreOffice(그래프 PPT → EMF 변환, UNO) + 글꼴 도구
RUN apt-get update && apt-get install -y --no-install-recommends \
        python3 python3-pip python3-uno libreoffice-impress fonts-noto-cjk fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt

# 보고서·그래프 글꼴 (KoPub돋움체, 함초롬돋움, 에스코어 드림) — EMF 글자 간격을 PowerPoint와 맞추는 데 필요
COPY fonts/ /usr/share/fonts/truetype/company/
RUN fc-cache -f

COPY . .
ENV DATA_DIR=/tmp/hwtx-sessions PYTHONUNBUFFERED=1
EXPOSE 8000
CMD ["sh", "-c", "python3 -m uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1 --timeout-keep-alive 120"]

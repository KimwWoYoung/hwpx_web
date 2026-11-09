#!/bin/sh
# 내 PC에서 실행: http://localhost:8000
pip install -r requirements.txt
export $(grep -v '^#' .env | xargs) 2>/dev/null
uvicorn app.main:app --port 8000

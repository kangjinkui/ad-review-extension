#!/bin/bash
# Linux/Mac 빌드 스크립트 (개발/테스트용)
echo "=== 광고물 심의 검토서 자동생성 - 빌드 ==="

echo "[1/2] 템플릿 파일 생성..."
python3 prepare_templates.py || { echo "템플릿 생성 실패"; exit 1; }

echo "[2/2] PyInstaller 빌드..."
pyinstaller \
    --onefile \
    --windowed \
    --name "광고물검토서자동생성" \
    --add-data "templates:templates" \
    app.py

if [ -f "dist/광고물검토서자동생성" ]; then
    echo "빌드 완료: dist/광고물검토서자동생성"
else
    echo "빌드 실패"
    exit 1
fi

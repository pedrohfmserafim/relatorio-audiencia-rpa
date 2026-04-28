#!/usr/bin/env bash
set -e

pip install -r requirements.txt

export PLAYWRIGHT_BROWSERS_PATH=/opt/render/project/src/.playwright-browsers
mkdir -p "$PLAYWRIGHT_BROWSERS_PATH"

echo "Instalando Chromium em: $PLAYWRIGHT_BROWSERS_PATH"
python -m playwright install chromium

echo "Conteúdo do diretório de browsers:"
ls "$PLAYWRIGHT_BROWSERS_PATH"

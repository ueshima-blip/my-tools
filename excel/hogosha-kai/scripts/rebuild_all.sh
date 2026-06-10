#!/usr/bin/env bash
# 保護者会調整_ver200.xlsm を「オリジナル + vba/ + XML手術」からクリーンに再構築する。
# 中間状態への上書きで壊れるのを避けるため、必ずオリジナルから作り直す。
#
#   bash scripts/rebuild_all.sh
#
set -euo pipefail
cd "$(dirname "$0")/.."

ORIG="保護者会調整_ver111_オリジナル.xlsm"
OUT="保護者会調整_ver200.xlsm"

echo "1/4 オリジナルから複製"
cp "$ORIG" "$OUT"

echo "2/4 VBA を注入 (vba/ → $OUT)"
python3 scripts/build_vba.py >/dev/null

echo "3/4 XML 手術 (ActiveX除去・ボタン・氏名1セル・時刻表示・calcChain削除ほか)"
python3 scripts/patch_workbook_xml.py

echo "4/4 構造検証"
python3 scripts/lint_vba.py >/dev/null
python3 scripts/validate_xlsx.py

echo "✓ 完了: $OUT"

#!/usr/bin/env bash
# 설치_한번에.js 를 Code.js + 대시보드.js + 설치.js 로 다시 만든다(편집기에 한 번에 붙여넣는 용도).
set -euo pipefail
REPO="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO"

OUT="$(mktemp)"
{
  cat <<'HEAD'
/**
 * 생산·품질 대시보드 — 설치_한번에 (편집기에 이 파일 하나만 붙여넣기)
 * ------------------------------------------------------------
 * Code.js + 대시보드.js + 설치.js 를 합친 파일입니다.
 * 붙여넣은 뒤 실행 순서:  설치_전체  →  설치_샘플데이터  →  설치_확인
 * (clasp 로 올릴 때는 이 파일을 올리지 않습니다 — 함수 이름이 겹칩니다.
 *  .claspignore 에서 제외해 두었습니다.)
 * ------------------------------------------------------------
 */

HEAD
  for f in Code.js 대시보드.js 설치.js; do
    printf '/* ===================== %s ===================== */\n\n' "$f"
    cat "$f"
    printf '\n'
  done
} > "$OUT"

mv "$OUT" 설치_한번에.js
echo "생성: 설치_한번에.js $(wc -c < 설치_한번에.js) bytes / $(wc -l < 설치_한번에.js) lines"
node --check 설치_한번에.js && echo "node --check: OK"
echo "--- 중복 선언 검사 ---"
DUP="$(grep -oE '^(function|const|let) [A-Za-z0-9_가-힣]+' 설치_한번에.js | sort | uniq -d || true)"
if [ -n "$DUP" ]; then echo "$DUP" | sed 's/^/  중복: /'; exit 1; else echo "  중복 없음"; fi
grep -c '^function ' 설치_한번에.js | sed 's/^/  함수 수: /'

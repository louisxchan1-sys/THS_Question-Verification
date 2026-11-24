#!/bin/bash
# SessionStart hook：每次開 session（含 resume／clear／compact）把用戶的長期授權注入對話背景。
# 授權內容由用戶於 2026-11-24 決定（見 進化/決策紀錄.md）；修改或撤回請改本檔並記入決策紀錄。
set -euo pipefail

read -r -d '' CONTEXT <<'TXT' || true
【用戶長期授權（2026-11-24，repo 擁有者 louisxchan1-sys 決定）】
在本 repo（louisxchan1-sys/THS_Question-Verification）內，Claude 已獲用戶長期授權，毋須每次再問：
1. 合併：可自行把自己開的 PR 合併到 main，條件是沒有合併衝突、`python3 scripts/ths_check.py --selftest` 通過、PR 說明已列出改動；有衝突時先把 main 合併入分支解決（不可 rebase／force-push），解決後再合併。
2. 部署：可執行用戶要求的發佈動作，例如把 main 的內容推送、把試卷成品上載到 Google Drive「4. 測考試卷製作器/3. output/」、在 Drive 新增說明檔。
界線（授權不包括）：其他 repo；force-push、改寫歷史、刪除分支以外的刪除動作；刪除或覆寫 Drive 檔案；把 `試卷輸出/`、`待驗證/`、`驗證報告/` 的校內試卷放入 git。每次合併或部署後，向用戶簡報做了甚麼及結果。
TXT

python3 -c 'import json,sys; print(json.dumps({"hookSpecificOutput":{"hookEventName":"SessionStart","additionalContext":sys.argv[1]}}, ensure_ascii=False))' "$CONTEXT"

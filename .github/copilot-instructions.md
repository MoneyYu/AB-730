# AB-730 repository conventions

- 使用繁體中文（台灣用語）；官方 module 名稱、產品名稱、指令與路徑保留原文。
- README 是 HackMD 學員頁。逐字保留 `## Links` 前的內容及 `## Exam` 到 EOF；日期、Course ID、survey、三語教材名稱／URL 不可自行更新。
- `## Links` 依 M01–M08 官方完整名稱與順序使用三級標題；首個連結緊接標題，獨立連結間一個空白行，不放 `/training/modules/` 或 unit URLs，不用 Azure 工程文件填充。
- Videos 只收 YouTube oEmbed 證實 LIVE／Microsoft 官方的內容，使用一張含 `Module` 欄的表；Mind Map 在原 `## Exam` 前，依八個 module 呈現課綱。
- 新連結逐一查 HTTP、final URL、title、locale 與語意；用既有 `.github\skills\course-prep\scripts\link_check.py`，不能只看 exit code。Evidence／link ledger 存本機 session，不提交。
- 文件驗證至少做 README baseline 比較、module 名稱／排序、禁止 URL、markmap fence／位置與 `git --no-pager diff --check`；文件-only 的 build／engine tests 為 N/A。
- Python 一律 `uv`；本機命令環境設定 `UV_INDEX_URL=https://packagefeedproxy.microsoft.io/pypi/simple/`，不裝全域套件、不改全域 registry。
- 受授權 PPT／speaker notes、既有 DEMO、未追蹤 `seed-data\` 只在本機讀取；不雲端上傳解析、不提交教材全文、帳號、tenant／app／resource IDs、secret、token 或 receipts。
- 講師情境、授權事實與實際未驗證差異放 `docs\`，不放 README。舊課程 seeding 成功不算 AB-730 驗證，template 不是認證；有效 config 缺少就維持 blocked。
- 本階段租戶盤點最多限定 GET，集合需分頁與上限；不寄信、上傳、建立／修補資源、擴權、登入其他角色或啟動 Cowork／排程。正式資料產生與加法灌入另需精確核准。
- 不覆寫既有未追蹤內容或使用者修改；以明確檔案 allowlist 操作，禁止 `git add -A`。本機素材的程式行號引用不代表 fresh clone 已包含 seeder。

# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## 常用指令

```bash
# 本地開發
python app.py                          # http://localhost:5000

# Deploy（GitHub push 自動 trigger Vercel）
git add -A && git commit -m "..." && git push
```

## 架構

```
用戶填表 (index.html) → POST /generate → generate_quotation() 生成 .xlsx
                    → 回傳 preview_id → 前端彈窗顯示下載連結
                    → GET /download/<pid>/excel → 下載 Excel
```

**核心流程**：純 Python，唔需要外部依賴。Flask 做 web server，openpyxl 生成 Excel。

### 關鍵檔案

| 檔案 | 角色 |
|------|------|
| `app.py` | Flask 路由：`/`(表單), `/generate`(生成), `/download/<pid>/excel` |
| `generator.py` | `generate_quotation(ws, data, title)` — 將 data dict 寫入 openpyxl worksheet，含 header、section、items、小計、總計、付款、條款 |
| `styles.py` | openpyxl 字型/填滿/邊框/對齊常數 + helper functions + A4 print setup |
| `templates/index.html` | 輸入表單：基本資料 + 兩個 textarea(基本/後加工程) + 自動分類 + 付款/條款設定 |
| `VERSION` | 版本號唯一來源，`app.py` 啟動時讀取，`index.html` 用 Jinja `{{ version }}` 渲染 |

### Data flow

**輸入 data dict 結構**：`{project_name, owner_name, address, company_name, quotation_no, date, validity, version, deposit, items: [{category, description, quantity, unit, unit_price, remark, is_additional}], payments: [{label, pct, desc}], terms: [str], show_payment, show_terms}`

**items 來源**：兩個 textarea 文字 → 前端 `parseItems()` 用 regex 解析 `{序號}{描述} {價錢}` → 前端關鍵字自動分類 → `collectData()` 打包 JSON → POST 到 `/generate`。

### 版本號規則

**只改 `VERSION` 一個檔案**。`app.py` 同 `index.html` 都係動態讀取，唔使手動 sync。每次 commit 必須確保 `VERSION` 內容正確。

## 注意

- 全部用廣東話/繁體中文溝通
- 呢個 project deploy 去 GitHub，**唔係** deploy 去 Vercel（雖然 vercel.json 存在）
- PDF/JPG 功能已暫時移除，目前只出 Excel

# 裝修報價單/發票助手

網頁版裝修報價單及發票生成工具，支援 Excel / PDF / JPG 輸出。

## 功能

- 網頁表單輸入工程項目（支援大字框批次輸入 + 自動分類）
- 生成專業格式報價單或發票（Excel .xlsx）
- 自動轉換 PDF（MiniPdf，格式與 Excel 完全一致）
- 自動轉換 JPG（PyMuPDF，高清輸出）
- 即時預覽（PDF 嵌入式預覽頁）
- 逆向匯入舊 Excel 檔案

## 安裝

```bash
pip install -r requirements.txt
```

### PDF 轉換（需要 MiniPdf）

```bash
# 安裝 .NET 9.0 SDK（如未安裝）
winget install Microsoft.DotNet.SDK.9

# 安裝 MiniPdf CLI
dotnet tool install --global MiniPdf.Cli
```

## 使用

```bash
python app.py
```

瀏覽器打開 `http://localhost:5000`

## 技術棧

- **後端**: Flask
- **Excel**: openpyxl
- **PDF**: MiniPdf (.NET)
- **JPG**: PyMuPDF
- **前端**: HTML/CSS/JavaScript

## 授權

MIT

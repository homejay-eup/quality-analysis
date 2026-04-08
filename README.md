# 設備品質分析 Web 版 — 建置指南

## 環境安裝
```bash
pip install streamlit pandas plotly openpyxl
```

## 檔案結構
```
專案資料夾/
├── app.py
├── requirements.txt
├── .gitignore
└── .streamlit/
    └── secrets.toml   # 本地用，不上傳 GitHub
```

## Excel 必要工作表與欄位

| 工作表 | 關鍵欄位 |
|--------|---------|
| 車機_彙整總覽 | 廠牌型號, ERP品號, 品名, 上線量, AB點, 失聯, 定位不良, 訊號異常, 其他, 回廠量, 回廠良品數, 回廠不良品數, 回廠過保數 |
| 鏡頭_彙整總覽 | 廠牌型號, ERP品號, 品名, 上線量, 黑畫面, 進水/模糊, 水波紋, 時有時無, 其他, 回廠量, 回廠良品數, 回廠不良品數, 回廠過保數 |

## 本地執行
```bash
cd 專案資料夾
python -m streamlit run app.py
# 開啟 http://localhost:8501
```

## 部署到 Streamlit Community Cloud

### 1. GitHub 初始化（首次）
```bash
git init
git add app.py requirements.txt .gitignore
git commit -m "Initial commit"
git remote add origin https://github.com/帳號/repo名稱.git
git branch -M main
git push -u origin main
```

### 2. 後續更新
```bash
git add app.py
git commit -m "更新說明"
git push
```
> Streamlit Cloud 偵測到 push 會自動重新部署

### 3. Streamlit Cloud 設定
1. [share.streamlit.io](https://share.streamlit.io) → New app
2. Repository: `帳號/repo名稱`　Branch: `main`　File: `app.py`
3. GitHub repo 須設為 **Public**，或至 `github.com/settings/installations` 授權 Streamlit 存取 Private repo
4. Deploy

### 4. Secrets 設定（有 API key 時）
- **雲端**：Streamlit Cloud → App → Settings → Secrets
- **本地**：`.streamlit/secrets.toml`（已被 .gitignore 排除）
```toml
[api]
key = "your-key"
```
程式碼內用 `st.secrets["api"]["key"]` 讀取

## 常見錯誤排查

| 錯誤 | 原因 | 解法 |
|------|------|------|
| `KeyError: '廠牌型號'` | 欄位名稱編碼問題 | `brand_col = next((c for c in ["廠牌型號","類型"] if c in df.columns), df.columns[0])` |
| `ImportError: matplotlib` | `background_gradient` 需要 matplotlib | 改用 `.format()` 不用 `.background_gradient()` |
| 鏡頭分析無資料 | `uploaded_file` stream 讀完第一張後位置到底 | `file_bytes = uploaded_file.read()` 再各自 `BytesIO(file_bytes)` |
| Streamlit 顯示 repo 不存在 | Repo 為 Private | 改 Public 或到 GitHub settings/installations 授權 |

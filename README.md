# 設備品質分析 Web 版 — 建置與維護指南

## 快速資訊
| 項目 | 內容 |
|------|------|
| 永久網址 | https://quality-analysis-qc6aehxkuqftgnizrkp66g.streamlit.app/ |
| GitHub | https://github.com/homejay-eup/quality-analysis |
| 主程式 | `互動式Web/app.py` |
| 部署方式 | push → Streamlit Cloud 自動重新部署 |

## 環境安裝
```bash
pip install streamlit pandas plotly openpyxl
```

## Excel 必要工作表與欄位
| 工作表 | 關鍵欄位 |
|--------|---------|
| 車機_彙整總覽 | 廠牌型號, ERP品號, 品名, 上線量, 回廠量, 回廠良品數, 回廠不良品數, 回廠過保數, AB點, 失聯, 定位不良, 訊號異常, 其他 |
| 鏡頭_彙整總覽 | 廠牌型號, ERP品號, 品名, 上線量, 回廠量, 回廠良品數, 回廠不良品數, 回廠過保數, 黑畫面, 進水/模糊, 水波紋, 時有時無, 其他 |

> 欄位名稱若不存在會 fallback 至明細欄加總（見 `process()` 函數）

---

## 現有功能清單（維護時勿破壞）

### ⚙️ 顯示設定（每 Tab 獨立，expander 預設收起，三欄佈局）

> **注意：已無 sidebar，所有設定移至各 Tab 內部的 `⚙️ 顯示設定` expander**

#### cfg1：長條圖配色 + 數值顯示模式
- 上線量顏色 `#4e79a7`、回廠量顏色 `#f28e2b`
- 不良率漸層低/高（`bad_scale`）、再使用率漸層低/高（`reuse_scale`）
- **🔢 數值顯示模式**：`st.radio`，4 選項 A/B/C/D
  - A=外部12px、B=auto12px、C=外部9px、D=不顯示
  - 套用至：不良率長條、再使用率長條（長條 texttemplate）、回廠原因/處置結果/上線量構成（圓環 textinfo）、上線量vs回廠量長條

#### cfg2：整體指標顯示 + 圖表顯示設定
- **📊 整體指標顯示**：7 個 checkbox（全預設勾選）
  - 總上線量、期間回廠量、期間再使用率、期間不良率、期間過保率、整體不良率、整體過保率
- **📈 圖表顯示設定**：6 個 checkbox

  | 圖表 | key | 預設 |
  |------|-----|------|
  | 各類型不良率長條圖 | `{name}_ch_bad` | **不顯示** |
  | 回廠原因分佈圓環圖 | `{name}_ch_fault` | 顯示 |
  | 各類型再使用率長條圖 | `{name}_ch_reuse` | **不顯示** |
  | 處置結果分佈圓環圖 | `{name}_ch_disp` | 顯示 |
  | 上線量 vs 回廠量長條圖 | `{name}_ch_onret` | **不顯示** |
  | 上線量構成圓環圖 | `{name}_ch_comp` | 顯示 |

#### cfg3：圓環圖配色
- **處置結果 / 上線量構成**（4 個共用色票）
  - 良品（再使用）`#4472c4`、不良品（維修/換貨）`#e74c3c`、過保/報廢 `#95a5a6`、仍在線上數 `#4e79a7`
- **回廠原因分佈**：依 fault_cols 動態生成色票（預設調色盤 `["#636efa","#ef553b","#00cc96","#ab63fa","#ffa15a"]`），車機/鏡頭各自獨立

---

### 篩選器（每個 Tab 獨立 key）
- **類型篩選**：`st.multiselect`，key=`{name}_brand`
- **ERP品號篩選**：顯示「品號 - 品名」複合標籤（`_erp_label` 欄），依類型篩選結果聯動縮減選項，key=`{name}_erp`

---

### 整體指標 KPI
共 7 格，順序固定，各自可由 cfg2 checkbox 控制顯示：

`總上線量` → `期間回廠量` → `期間再使用率` → `期間不良率` → `期間過保率` → `整體不良率` → `整體過保率`

| 指標 | 公式 |
|------|------|
| 期間再使用率 | 良品數 / 回廠量 |
| 期間不良率   | 不良品數 / 回廠量 |
| 期間過保率   | 過保數 / 回廠量 |
| 整體不良率   | 不良品數 / 上線量 |
| 整體過保率   | 過保數 / 上線量 |

> **kpi_row(df, show_flags)** 接受 show_flags dict，動態決定顯示欄數並呼叫 `st.columns(n)`

---

### 圖表（每 Tab 共 6 張，動態兩欄配對佈局）

可見圖表依序兩兩配對為 `st.columns(2)`，未勾選的圖表不渲染也不佔位。

| 函數 | 圖型 | 說明 |
|------|------|------|
| `make_bar_bad` | 長條 | 各類型整體不良率（整體不良率(%)=不良品數/上線量），漸層=bad_scale |
| `make_pie_fault` | 圓環 | 回廠原因分佈（fault_cols），hole=0.45 |
| `make_bar_reuse` | 長條 | 各類型再使用率，漸層=reuse_scale |
| `make_pie_disp` | 圓環 | 處置結果分佈（良品/不良品/過保），hole=0.45 |
| `make_bar_onret` | 分組長條 | 上線量 vs 回廠量，go.Figure 手刻 |
| `make_pie_comp` | 圓環 | 上線量構成（良品/不良品/過保/仍在線上數），hole=0.45 |

> **Y軸標題處理**：make_bar_bad / make_bar_reuse 用 `yaxis_title=""` + `annotations`（x=-0.07, y=1.08），勿改回 `yaxis_title=` 否則會恢復垂直

---

### process() 計算欄位

| 欄位 | 公式 | 說明 |
|------|------|------|
| 再使用率(%) | 良品數 / 回廠量 × 100 | 期間效率 |
| 不良率(%)   | 不良品數 / 回廠量 × 100 | 期間不良 |
| 過保率(%)   | 過保數 / 回廠量 × 100 | 期間過保 |
| 整體不良率(%) | 不良品數 / 上線量 × 100 | 整體不良 |
| 整體過保率(%) | 過保數 / 上線量 × 100 | 整體過保 |

---

### 詳細資料表
- **欄位順序**：類型, ERP品號, 品名, 上線量, 回廠量, 良品數, 再使用率(%), 不良品數, 不良率(%), 過保數, 過保率(%), 整體不良率(%), 整體過保率(%)
- 原始 `brand_col` 欄在顯示時 rename 為「類型」
- 數量欄（上線量/回廠量/良品數/不良品數/過保數）全部 `astype(int)`，格式 `{:,}`
- 比率欄格式 `{:.1f}`（共 5 欄）
- 小計列：groupby 後重新計算所有比率欄；樣式 `background-color: #d6e4f0; font-weight: bold`
- 總計列：全體加總後重新計算所有比率欄；樣式 `background-color: #1f4e79; color: white; font-weight: bold`
- **摺疊開關**：`st.toggle`，key=`{name}_collapse`，開啟時僅顯示小計+總計，隱藏 ERP品號/品名欄，表格高度 `max(200, min(60*n+38, 600))`

---

## 程式碼關鍵結構

```
uploaded_file → file_bytes（BytesIO 複用）
  process(df, fault_cols)
    → 標準化欄位
    → 計算：再使用率(%), 不良率(%), 過保率(%),（均除以回廠量）
            整體不良率(%), 整體過保率(%)（均除以上線量）
  kpi_row(df, show_flags) → 最多 7 格 metric，動態 st.columns(n)
  render_tab(df, name, fault_cols)
    brand_col 動態偵測（廠牌型號/廠牌/類型/第一欄）
    _erp_label 複合欄（品號 - 品名）
    ⚙️ 顯示設定 expander（cfg1/cfg2/cfg3 三欄）
    篩選：brand_filtered → filtered
    kpi_row(filtered, show_kpi)
    make_bar_bad / make_pie_fault / make_bar_reuse /
    make_pie_disp / make_bar_onret / make_pie_comp
      → 動態兩欄配對，僅渲染勾選的圖表
    詳細資料表（subtotals groupby → frames concat → rename → toggle摺疊）
    download_button
  tab1（車機）、tab2（鏡頭）各呼叫 render_tab
```

---

## 更新部署
```bash
cd "C:\Users\EUser\Desktop\品質分析報告v2\互動式Web"
git add app.py
git commit -m "說明"
git push origin main
# 若 push 被拒（remote 有新 commit）：
git pull --rebase origin main
git push origin main
```

---

## 常見錯誤排查
| 錯誤 | 原因 | 解法 |
|------|------|------|
| `KeyError: '廠牌型號'` | 欄位名稱不符 | `brand_col = next((c for c in ["廠牌型號","廠牌","類型"] if c in df.columns), df.columns[0])` |
| 鏡頭分析無資料 | stream 讀完位置到底 | `file_bytes = uploaded_file.read()` 各自 `BytesIO(file_bytes)` |
| ERP聯動篩選失效 | erp_options 來源用了原始 df | 確認 erp_options 來自 `brand_filtered` 而非 `df` |
| 數量欄出現小數 | groupby sum 產生 float | `astype(int)` 須在 rename 之後、styled 之前執行 |
| Y軸標題變垂直 | 誤設 `yaxis_title=文字` | 保持 `yaxis_title=""` + annotation 方式 |
| push 被拒 | remote 有新 commit | `git pull --rebase origin main` 再 push |
| 圓環圖顏色未套用 | `color_discrete_map` key 與 names 不符 | 確認 map 的 key 字串與 `px.pie(names=...)` 完全一致 |
| widget key 衝突 | 兩 Tab 共用同一 key | 所有 widget 均加 `key=f"{name}_..."` 前綴 |

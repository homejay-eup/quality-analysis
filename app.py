import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="設備品質分析", layout="wide", page_icon="📊")

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 2rem; }
[data-testid="stMetricLabel"] { font-size: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("📊 設備品質分析 Web 版")

# ── 常數 ─────────────────────────────────────────────────────────────────────
CAR_FAULT_COLS  = ["AB點", "失聯", "定位不良", "訊號異常", "其他"]
LENS_FAULT_COLS = ["黑畫面", "進水/模糊", "水波紋", "時有時無", "其他"]

# ── 上傳 ─────────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "上傳設備品質分析 Excel（多期合併版）",
    type=["xlsx"],
    help="需含：車機_彙整總覽、鏡頭_彙整總覽（含「期間」欄）；選填：車機_月趨勢、鏡頭_月趨勢",
)

if not uploaded_file:
    st.info(
        "請上傳多期合併版分析總表。\n\n"
        "必要工作表：**車機_彙整總覽**、**鏡頭_彙整總覽**（需含 `期間` 欄，如 `2025-Q1`、`2026-Q1`）\n\n"
        "選填工作表：**車機_月趨勢**、**鏡頭_月趨勢**（含 `期間`、`年月`、`類型` 及各指標欄）"
    )
    st.stop()

file_bytes = uploaded_file.read()

@st.cache_data(show_spinner="讀取資料中…")
def load_sheets(fb):
    xls = pd.ExcelFile(BytesIO(fb))
    return {sh: xls.parse(sh) for sh in xls.sheet_names}

try:
    sheets = load_sheets(file_bytes)
except Exception as e:
    st.error(f"無法讀取檔案：{e}")
    st.stop()

def get_sheet(name):
    return sheets.get(name, pd.DataFrame()).copy()

# ── Y 軸垂直文字（annotation 方式，每字換行）─────────────────────────────────
def apply_vlabel(fig, text, margin_l=75):
    """以 annotation 取代 y-axis title，實現字元逐行堆疊的垂直中文標籤。"""
    stacked = "<br>".join(list(str(text)))
    fig.update_yaxes(title_text="")
    fig.update_layout(margin={"l": margin_l})
    fig.add_annotation(
        text=stacked,
        showarrow=False,
        xref="paper",
        yref="paper",
        x=-0.045,
        y=0.5,
        xanchor="center",
        yanchor="middle",
        font=dict(size=12),
    )

# ── 資料前處理 ────────────────────────────────────────────────────────────────
def process(df, fault_cols):
    df = df.copy()
    df.columns = df.columns.str.strip()
    if "期間" not in df.columns:
        df["期間"] = "未知期間"

    good_col  = next((c for c in ["回廠良品數",  "良品數"]  if c in df.columns), None)
    bad_col   = next((c for c in ["回廠不良品數", "不良品數"] if c in df.columns), None)
    scrap_col = next((c for c in ["回廠過保數",  "過保數"]  if c in df.columns), None)

    if good_col is None:
        df["回廠良品數"] = df[[c for c in ["O /測試正常","回廠QC","回廠其他"] if c in df.columns]].sum(axis=1)
        good_col = "回廠良品數"
    if bad_col is None:
        df["回廠不良品數"] = df[[c for c in ["G /評估後退修","X /已完修","維修換貨＋換貨條碼"] if c in df.columns]].sum(axis=1)
        bad_col = "回廠不良品數"
    if scrap_col is None:
        df["回廠過保數"] = df[[c for c in ["D /停產報廢","E /過保報廢","回廠報廢"] if c in df.columns]].sum(axis=1)
        scrap_col = "回廠過保數"
    if "回廠量" not in df.columns:
        df["回廠量"] = df[[c for c in fault_cols if c in df.columns]].sum(axis=1)

    df["良品數"]        = df[good_col].fillna(0)
    df["不良品數"]      = df[bad_col].fillna(0)
    df["過保數"]        = df[scrap_col].fillna(0)
    df["再使用率(%)"]   = (df["良品數"]   / df["回廠量"] * 100).fillna(0).round(1)
    df["不良率(%)"]     = (df["不良品數"] / df["回廠量"] * 100).fillna(0).round(1)
    df["過保率(%)"]     = (df["過保數"]   / df["回廠量"] * 100).fillna(0).round(1)
    df["整體不良率(%)"] = (df["不良品數"] / df["上線量"] * 100).fillna(0).round(1)
    df["整體過保率(%)"] = (df["過保數"]   / df["上線量"] * 100).fillna(0).round(1)
    return df

# ── KPI 計算 ──────────────────────────────────────────────────────────────────
def calc_kpi(df):
    if df is None or len(df) == 0:
        return None
    ton  = int(df["上線量"].sum())
    tret = int(df["回廠量"].sum())
    tbad = int(df["不良品數"].sum())
    tgd  = int(df["良品數"].sum())
    tscr = int(df["過保數"].sum()) if "過保數" in df.columns else 0
    return dict(
        total_on=ton, total_ret=tret, total_bad=tbad, total_good=tgd, total_scrap=tscr,
        reuse_rate   = tgd  / tret * 100 if tret else 0,
        bad_rate     = tbad / tret * 100 if tret else 0,
        scrap_rate   = tscr / tret * 100 if tret else 0,
        ovr_bad_rate = tbad / ton  * 100 if ton  else 0,
        ovr_scr_rate = tscr / ton  * 100 if ton  else 0,
    )

# ── 品管建議生成 ──────────────────────────────────────────────────────────────
def gen_quality_rec(kpi_c, kpi_p, pc, pp, name, fault_cols, df_c):
    pp_label = pp if kpi_p else "（無對比資料）"
    lines = [f"【{name} 品管建議】　{pc} vs {pp_label}\n" + "─"*40]

    def delta(c, p, key):
        return (c[key] - p[key]) if p else None

    bad = kpi_c["ovr_bad_rate"]
    d   = delta(kpi_c, kpi_p, "ovr_bad_rate")
    d_s = f"（較去年同期 {'+' if d and d>=0 else ''}{d:.1f}%）" if d is not None else ""
    if bad >= 5:
        lines.append(f"⚠️ 整體不良率 {bad:.1f}% 偏高{d_s}，建議加強來料 IQC 抽驗標準，並要求供應商提交 8D 改善報告。")
    elif d is not None and d > 1:
        lines.append(f"📈 整體不良率 {bad:.1f}%{d_s}，呈上升趨勢，請追蹤是否有批次性品質異常。")
    else:
        lines.append(f"✅ 整體不良率 {bad:.1f}%{d_s}，品質表現穩定。")

    reuse = kpi_c["reuse_rate"]
    dr    = delta(kpi_c, kpi_p, "reuse_rate")
    dr_s  = f"（較去年同期 {'+' if dr and dr>=0 else ''}{dr:.1f}%）" if dr is not None else ""
    if reuse >= 70:
        lines.append(f"✅ 再使用率 {reuse:.1f}%{dr_s}，維修品質良好，可維持現行維修 SOP。")
    else:
        lines.append(f"⚠️ 再使用率 {reuse:.1f}%{dr_s}，偏低，建議檢視維修流程及技師技術水準，必要時安排教育訓練。")

    avail = [c for c in fault_cols if c in df_c.columns]
    if avail:
        ft = df_c[avail].sum().sort_values(ascending=False)
        ft = ft[ft > 0]
        if not ft.empty:
            top, top_pct = ft.index[0], ft.iloc[0] / ft.sum() * 100
            if top_pct >= 40:
                lines.append(f"🔍 主要故障原因為「{top}」（佔 {top_pct:.1f}%），建議針對此問題建立專項改善計畫，設定月度追蹤指標。")
            else:
                top3 = "、".join(f"{i}（{ft[i]/ft.sum()*100:.0f}%）" for i in ft.index[:3])
                lines.append(f"🔍 故障原因較分散，前三項為：{top3}，建議各項分別追蹤改善。")

    scr = kpi_c["ovr_scr_rate"]
    ds  = delta(kpi_c, kpi_p, "ovr_scr_rate")
    ds_s = f"（較去年同期 {'+' if ds and ds>=0 else ''}{ds:.1f}%）" if ds is not None else ""
    if scr >= 3:
        lines.append(f"⚠️ 整體過保率 {scr:.1f}%{ds_s}，建議統計過保品平均使用年限，作為下次汰換計畫依據，並檢視保固條款是否需調整。")
    else:
        lines.append(f"✅ 整體過保率 {scr:.1f}%{ds_s}，目前過保狀況在可控範圍內。")

    lines.append("\n*以上建議為系統自動生成，請依實際情況修改後使用。*")
    return "\n".join(lines)

# ── 採購建議生成 ──────────────────────────────────────────────────────────────
def gen_purchase_rec(kpi_c, kpi_p, pc, pp, name, df_c):
    pp_label = pp if kpi_p else "（無對比資料）"
    lines = [f"【{name} 採購建議】　{pc} vs {pp_label}\n" + "─"*40]
    brand_col = next((c for c in ["廠牌型號","廠牌","類型"] if c in df_c.columns), df_c.columns[0])

    if "已使用年限" in df_c.columns and "過保率(%)" in df_c.columns:
        hi = df_c[(df_c["過保率(%)"] >= 10) & (df_c["已使用年限"].notna())]
        if not hi.empty:
            items = hi.nlargest(3, "過保率(%)")[brand_col].tolist()
            lines.append(f"🔄 下列類型過保率偏高（≥10%），建議優先列入汰換採購計畫：\n   {'、'.join(map(str,items))}")
        else:
            lines.append("✅ 目前無高過保率品項需緊急汰換。")
    else:
        lines.append("（需有「已使用年限」及「過保率(%)」欄位，方可自動判斷汰換建議）")

    if "整體不良率(%)" in df_c.columns:
        hi2 = df_c[df_c["整體不良率(%)"] >= 5]
        if not hi2.empty:
            items2 = hi2.nlargest(3, "整體不良率(%)")[brand_col].tolist()
            lines.append(f"📋 下列類型整體不良率偏高（≥5%），建議要求廠商提出品質改善計畫，或於下次採購時評估替換供應商：\n   {'、'.join(map(str,items2))}")

    if kpi_p and kpi_p["total_ret"] > 0:
        delta_pct = (kpi_c["total_ret"] - kpi_p["total_ret"]) / kpi_p["total_ret"] * 100
        if delta_pct >= 15:
            lines.append(f"📈 回廠量較去年同期上升 {delta_pct:.1f}%，建議採購合約加入品質保證條款，並提高新批次 AQL 抽樣比例。")
        elif delta_pct <= -10:
            lines.append(f"📉 回廠量較去年同期下降 {abs(delta_pct):.1f}%，品質改善明顯，可維持現行採購策略，適時與供應商確認品質保持計畫。")

    lines.append("\n*以上建議為系統自動生成，請依實際情況修改後使用。*")
    return "\n".join(lines)

# ── 主渲染函數 ────────────────────────────────────────────────────────────────
def render_tab(sheet_name, trend_sheet_name, name, fault_cols):
    df_all = get_sheet(sheet_name)
    if df_all.empty:
        st.error(f"找不到工作表：{sheet_name}")
        return

    df_all = process(df_all, fault_cols)
    all_periods = sorted(df_all["期間"].dropna().unique().tolist(), reverse=True)

    pc1, pc2, _ = st.columns([2, 2, 4])
    with pc1:
        period_cur = st.selectbox("目前期間", all_periods, index=0, key=f"{name}_cur")
    with pc2:
        prv_opts = [p for p in all_periods if p != period_cur]
        period_prv = st.selectbox(
            "對比期間（去年同期）",
            prv_opts if prv_opts else ["（尚無對比資料）"],
            index=0, key=f"{name}_prv"
        )
    has_prv = bool(prv_opts)

    df_cur = df_all[df_all["期間"] == period_cur].copy()
    df_prv = df_all[df_all["期間"] == period_prv].copy() if has_prv else None

    brand_col  = next((c for c in ["廠牌型號","廠牌","類型"] if c in df_cur.columns), df_cur.columns[0])
    has_vendor = "廠商" in df_cur.columns
    has_erp    = "ERP品號" in df_cur.columns
    has_name   = "品名" in df_cur.columns

    if has_erp and has_name:
        def _make_label(df):
            df["_label"] = df.apply(
                lambda r: f"{str(r['ERP品號']).strip()} - {str(r['品名']).strip()}"
                if pd.notna(r["ERP品號"]) else str(r["品名"]), axis=1)
        _make_label(df_cur)
        if df_prv is not None:
            _make_label(df_prv)

    # ── 篩選器：廠商 → 類型 → ERP品號（階層聯動）
    if has_vendor:
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            sel_vendors = st.multiselect(
                "廠商篩選",
                sorted(df_cur["廠商"].dropna().unique().tolist()),
                key=f"{name}_vendor"
            )
        _v_df = df_cur[df_cur["廠商"].isin(sel_vendors)] if sel_vendors else df_cur
        with fc2:
            sel_brands = st.multiselect(
                "類型篩選",
                sorted(_v_df[brand_col].dropna().unique().tolist()),
                key=f"{name}_brand"
            )
        _b_df = _v_df[_v_df[brand_col].isin(sel_brands)] if sel_brands else _v_df
        with fc3:
            if has_erp:
                _erp_opts = (
                    _b_df["_label"].dropna().unique().tolist()
                    if "_label" in _b_df.columns
                    else _b_df["ERP品號"].dropna().unique().tolist()
                )
                sel_erp = st.multiselect("ERP品號篩選", _erp_opts, key=f"{name}_erp")
            else:
                sel_erp = []
                st.multiselect("ERP品號篩選", [], key=f"{name}_erp")
    else:
        fc1, fc2 = st.columns(2)
        sel_vendors = []
        with fc1:
            sel_brands = st.multiselect(
                "類型篩選",
                df_cur[brand_col].dropna().unique(),
                key=f"{name}_brand"
            )
        with fc2:
            if has_erp:
                _fd = df_cur[df_cur[brand_col].isin(sel_brands)] if sel_brands else df_cur
                _erp_opts = (
                    _fd["_label"].dropna().unique().tolist()
                    if "_label" in _fd.columns
                    else _fd["ERP品號"].dropna().unique().tolist()
                )
                sel_erp = st.multiselect("ERP品號篩選", _erp_opts, key=f"{name}_erp")
            else:
                sel_erp = []
                st.multiselect("ERP品號篩選", [], key=f"{name}_erp")

    # 套用篩選
    if has_vendor and sel_vendors:
        df_cur = df_cur[df_cur["廠商"].isin(sel_vendors)]
        if df_prv is not None:
            df_prv = df_prv[df_prv["廠商"].isin(sel_vendors)]
    if sel_brands:
        df_cur = df_cur[df_cur[brand_col].isin(sel_brands)]
        if df_prv is not None:
            df_prv = df_prv[df_prv[brand_col].isin(sel_brands)]
    if sel_erp and "_label" in df_cur.columns:
        df_cur = df_cur[df_cur["_label"].isin(sel_erp)]
        if df_prv is not None and "_label" in df_prv.columns:
            df_prv = df_prv[df_prv["_label"].isin(sel_erp)]

    kpi_c = calc_kpi(df_cur)
    kpi_p = calc_kpi(df_prv) if df_prv is not None and len(df_prv) > 0 else None

    # ── 整體指標（彈性顯示）
    ALL_KPIS     = ["總上線量", "期間回廠量", "期間再使用率", "期間不良率", "整體不良率", "整體過保率"]
    DEFAULT_KPIS = ["總上線量", "期間回廠量", "期間再使用率", "整體不良率", "整體過保率"]

    kpi_hdr, kpi_cfg = st.columns([4, 2])
    kpi_hdr.markdown("#### 整體指標")
    with kpi_cfg:
        with st.expander("⚙️ 選擇顯示指標"):
            shown_kpis = st.multiselect(
                "顯示的整體指標", ALL_KPIS, default=DEFAULT_KPIS,
                key=f"{name}_kpi_sel", label_visibility="collapsed"
            )
    if not shown_kpis:
        shown_kpis = DEFAULT_KPIS

    def d_str(cur_v, prv_v, fmt=",.0f", suffix=""):
        if prv_v is None:
            return None
        d = cur_v - prv_v
        return f"{'+' if d >= 0 else ''}{d:{fmt}}{suffix}"

    kpis_all = [
        ("總上線量",     f"{kpi_c['total_on']:,}",        d_str(kpi_c['total_on'],      kpi_p['total_on']      if kpi_p else None)),
        ("期間回廠量",   f"{kpi_c['total_ret']:,}",       d_str(kpi_c['total_ret'],     kpi_p['total_ret']     if kpi_p else None)),
        ("期間再使用率", f"{kpi_c['reuse_rate']:.1f}%",   d_str(kpi_c['reuse_rate'],    kpi_p['reuse_rate']    if kpi_p else None, fmt=".1f", suffix="%")),
        ("期間不良率",   f"{kpi_c['bad_rate']:.1f}%",     d_str(kpi_c['bad_rate'],      kpi_p['bad_rate']      if kpi_p else None, fmt=".1f", suffix="%")),
        ("整體不良率",   f"{kpi_c['ovr_bad_rate']:.1f}%", d_str(kpi_c['ovr_bad_rate'],  kpi_p['ovr_bad_rate']  if kpi_p else None, fmt=".1f", suffix="%")),
        ("整體過保率",   f"{kpi_c['ovr_scr_rate']:.1f}%", d_str(kpi_c['ovr_scr_rate'],  kpi_p['ovr_scr_rate']  if kpi_p else None, fmt=".1f", suffix="%")),
    ]
    kpis = [(lbl, val, dlt) for lbl, val, dlt in kpis_all if lbl in shown_kpis]
    for i in range(0, len(kpis), 3):
        row_cols = st.columns(3)
        for j, col in enumerate(row_cols):
            if i + j < len(kpis):
                lbl, val, dlt = kpis[i + j]
                col.metric(lbl, val, dlt)

    st.markdown("---")

    # ── 趨勢 & 圖表：統一彈性顯示控制 ───────────────────────────────────────
    df_trend_raw = get_sheet(trend_sheet_name)
    trend_available = not df_trend_raw.empty

    trend_opt         = ["月趨勢"] if trend_available else []
    base_chart_opts   = ["回廠原因分佈", "處置結果分佈", "各類型不良率", "同期回廠量比較"]
    vendor_chart_opts = ["廠商同期比較", "廠商各指標分析"] if has_vendor else []
    all_display_opts  = trend_opt + base_chart_opts + vendor_chart_opts

    sec_hdr, sec_cfg = st.columns([4, 2])
    sec_hdr.markdown("#### 趨勢與圖表")
    with sec_cfg:
        with st.expander("⚙️ 選擇顯示內容"):
            sel_display = st.multiselect(
                "顯示內容", all_display_opts,
                default=trend_opt + base_chart_opts,
                key=f"{name}_display",
                label_visibility="collapsed"
            )
    if not sel_display:
        sel_display = trend_opt + base_chart_opts

    # ── 月趨勢（廠商→類型→ERP品號 階層篩選）
    if "月趨勢" in sel_display and trend_available:
        st.markdown("##### 📈 月趨勢")
        df_trend_raw.columns = df_trend_raw.columns.str.strip()
        trend_cols_avail = [c for c in ["回廠量","不良品數","良品數","過保數","上線量"] if c in df_trend_raw.columns]

        trend_periods = [period_cur] + ([period_prv] if has_prv else [])
        dt_base = df_trend_raw[df_trend_raw["期間"].isin(trend_periods)].copy()

        # 補充廠商資訊
        t_brand_col = next((c for c in ["廠牌型號","廠牌","類型"] if c in dt_base.columns), None)
        if has_vendor and t_brand_col:
            _join_col = brand_col if brand_col in dt_base.columns else t_brand_col
            _vmap = df_all[["廠商", brand_col]].drop_duplicates().rename(columns={brand_col: _join_col})
            if _join_col in dt_base.columns:
                dt_base = dt_base.merge(_vmap, on=_join_col, how="left")

        # 補充 ERP品號 - 品名 對照標籤
        t_has_erp  = "ERP品號" in dt_base.columns
        if t_has_erp and has_name:
            _en_map = (
                df_all[["ERP品號","品名"]]
                .dropna(subset=["ERP品號"])
                .drop_duplicates(subset=["ERP品號"])
            )
            dt_base = dt_base.merge(_en_map, on="ERP品號", how="left")
            dt_base["_t_label"] = dt_base.apply(
                lambda r: f"{str(r['ERP品號']).strip()} - {str(r['品名']).strip()}"
                if pd.notna(r.get("品名")) and str(r.get("品名")).strip() else str(r["ERP品號"]),
                axis=1
            )
        t_has_label  = "_t_label" in dt_base.columns
        t_has_vendor = "廠商" in dt_base.columns

        # 趨勢篩選器欄位數
        _tcols = 1 + (1 if t_has_vendor else 0) + (1 if t_brand_col else 0) + (1 if t_has_erp else 0)
        trend_ui = st.columns(_tcols)

        with trend_ui[0]:
            sel_metric = st.selectbox("趨勢指標", trend_cols_avail, key=f"{name}_tm")

        _tidx = 1
        t_sel_vendor = []
        t_sel_brand  = []
        t_sel_erp    = []

        if t_has_vendor:
            with trend_ui[_tidx]:
                t_sel_vendor = st.multiselect(
                    "廠商", sorted(dt_base["廠商"].dropna().unique().tolist()),
                    key=f"{name}_t_vendor"
                )
            _tidx += 1

        _dt_v = dt_base[dt_base["廠商"].isin(t_sel_vendor)] if t_sel_vendor else dt_base

        if t_brand_col:
            with trend_ui[_tidx]:
                t_sel_brand = st.multiselect(
                    "類型", sorted(_dt_v[t_brand_col].dropna().unique().tolist()),
                    key=f"{name}_t_brand"
                )
            _tidx += 1

        _dt_b = _dt_v[_dt_v[t_brand_col].isin(t_sel_brand)] if t_sel_brand and t_brand_col else _dt_v

        if t_has_erp:
            with trend_ui[_tidx]:
                _erp_label_col = "_t_label" if t_has_label else "ERP品號"
                t_erp_opts = sorted(_dt_b[_erp_label_col].dropna().unique().tolist())
                t_sel_erp = st.multiselect("ERP品號", t_erp_opts, key=f"{name}_t_erp")

        # 套用趨勢篩選
        dt = dt_base.copy()
        if t_sel_vendor and t_has_vendor:
            dt = dt[dt["廠商"].isin(t_sel_vendor)]
        if t_sel_brand and t_brand_col:
            dt = dt[dt[t_brand_col].isin(t_sel_brand)]
        if t_sel_erp and t_has_erp:
            _filter_col = "_t_label" if t_has_label else "ERP品號"
            dt = dt[dt[_filter_col].isin(t_sel_erp)]

        if "年月" in dt.columns and sel_metric in dt.columns:
            dt_agg = dt.groupby(["期間","年月"])[sel_metric].sum().reset_index()
            dt_agg["月份"] = dt_agg["年月"].astype(str).str[-2:] + "月"
            fig_t = px.line(
                dt_agg, x="月份", y=sel_metric, color="期間",
                markers=True,
                title=f"{name}｜{sel_metric} 月趨勢",
                color_discrete_map={period_cur:"#4e79a7", period_prv:"#f28e2b"},
            )
            fig_t.update_traces(line=dict(width=2.5), marker=dict(size=8))
            fig_t.update_layout(height=360, legend_title="期間")
            apply_vlabel(fig_t, sel_metric)
            st.plotly_chart(fig_t, use_container_width=True)
        else:
            st.info("月趨勢工作表缺少 `年月` 或對應指標欄位")

        st.markdown("---")

    elif "月趨勢" not in sel_display and not trend_available:
        with st.expander("💡 如何啟用月趨勢圖？", expanded=False):
            st.info(
                f"在 Excel 中新增「{trend_sheet_name}」工作表，欄位如下：\n\n"
                "`期間` | `年月` | `類型` | `上線量` | `回廠量` | `不良品數` | `良品數` | `過保數`\n\n"
                "例：`2026-Q1` | `2026-01` | `16-1車機` | 120 | 5 | 2 | 3 | 0"
            )

    # ── 圖表（由 sel_display 控制）
    charts_to_render = [c for c in sel_display if c != "月趨勢"]
    if charts_to_render:
        st.markdown("##### 📊 圖表分析")

    if "回廠原因分佈" in charts_to_render:
        st.markdown("###### 回廠原因分佈")
        left, right = st.columns(2)
        for col_w, df_w, label in [(left, df_cur, period_cur), (right, df_prv, period_prv if has_prv else None)]:
            if df_w is None or len(df_w) == 0:
                if label:
                    col_w.info(f"{label}：無資料")
                continue
            avail = [c for c in fault_cols if c in df_w.columns]
            totals = df_w[avail].sum()
            totals = totals[totals > 0]
            if not totals.empty:
                fig = px.pie(values=totals.values, names=totals.index, hole=0.45,
                             title=f"{name}｜回廠原因（{label}）")
                fig.update_traces(textinfo="percent+label", textposition="outside")
                fig.update_layout(height=400)
                col_w.plotly_chart(fig, use_container_width=True)

    if "處置結果分佈" in charts_to_render:
        st.markdown("###### 處置結果分佈")
        left, right = st.columns(2)
        for col_w, kpi_w, label in [(left, kpi_c, period_cur), (right, kpi_p, period_prv if has_prv else None)]:
            if kpi_w is None:
                if label:
                    col_w.info(f"{label}：無資料")
                continue
            disp = {"良品（再使用）": kpi_w["total_good"],
                    "不良品（維修/換貨）": kpi_w["total_bad"],
                    "過保/報廢": kpi_w["total_scrap"]}
            disp = {k: v for k, v in disp.items() if v > 0}
            if disp:
                _cm = {"良品（再使用）":"#4472c4","不良品（維修/換貨）":"#e74c3c","過保/報廢":"#95a5a6"}
                fig = px.pie(values=list(disp.values()), names=list(disp.keys()),
                             hole=0.45, color_discrete_map=_cm,
                             title=f"{name}｜處置結果（{label}）")
                fig.update_traces(textinfo="percent+label", textposition="outside")
                fig.update_layout(height=400)
                col_w.plotly_chart(fig, use_container_width=True)

    if "各類型不良率" in charts_to_render:
        st.markdown("###### 各類型不良率")
        fig = px.bar(
            df_cur.sort_values("整體不良率(%)", ascending=False),
            x=brand_col, y="整體不良率(%)",
            color="整體不良率(%)", color_continuous_scale=[[0,"#00b050"],[1,"#ff0000"]],
            title=f"{name}｜各類型整體不良率（{period_cur}）", text="整體不良率(%)",
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", textfont_size=12)
        fig.update_layout(coloraxis_showscale=False, height=440)
        apply_vlabel(fig, "整體不良率 (%)")
        st.plotly_chart(fig, use_container_width=True)

    if "同期回廠量比較" in charts_to_render:
        st.markdown("###### 同期回廠量比較")
        if df_prv is not None and len(df_prv) > 0:
            _c = df_cur[[brand_col,"回廠量"]].copy(); _c["期間"] = period_cur
            _p = df_prv[[brand_col,"回廠量"]].copy(); _p["期間"] = period_prv
            fig = px.bar(pd.concat([_c, _p]), x=brand_col, y="回廠量", color="期間", barmode="group",
                         title=f"{name}｜回廠量同期比較",
                         color_discrete_map={period_cur:"#4e79a7", period_prv:"#f28e2b"})
            fig.update_layout(height=440)
            apply_vlabel(fig, "回廠量")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("尚無對比期間資料，待匯入去年同期資料後自動顯示。")

    if "廠商同期比較" in charts_to_render and has_vendor:
        st.markdown("###### 廠商同期比較")
        _nm = [c for c in ["回廠量","不良品數","良品數","過保數"] if c in df_cur.columns]
        _cv = df_cur[["廠商"]+_nm].groupby("廠商")[_nm].sum().reset_index(); _cv["期間"] = period_cur
        combined_v = _cv.copy()
        if df_prv is not None and len(df_prv) > 0:
            _pv = df_prv[["廠商"]+_nm].groupby("廠商")[_nm].sum().reset_index(); _pv["期間"] = period_prv
            combined_v = pd.concat([_cv, _pv], ignore_index=True)
        fig = px.bar(combined_v, x="廠商", y="回廠量", color="期間", barmode="group",
                     title=f"{name}｜廠商回廠量同期比較",
                     color_discrete_map={period_cur:"#4e79a7", period_prv:"#f28e2b"})
        fig.update_layout(height=440)
        apply_vlabel(fig, "回廠量")
        st.plotly_chart(fig, use_container_width=True)

    if "廠商各指標分析" in charts_to_render and has_vendor:
        st.markdown("###### 廠商各指標分析")
        _avail_m = [c for c in ["回廠量","不良品數","良品數","過保數"] if c in df_cur.columns]
        vendor_metrics = st.multiselect(
            "選擇廠商分析指標", _avail_m,
            default=_avail_m[:2] if len(_avail_m) >= 2 else _avail_m,
            key=f"{name}_vmetrics"
        )
        _cva = df_cur[["廠商"]+_avail_m].groupby("廠商")[_avail_m].sum().reset_index(); _cva["期間"] = period_cur
        combined_va = _cva.copy()
        if df_prv is not None and len(df_prv) > 0:
            _pva = df_prv[["廠商"]+_avail_m].groupby("廠商")[_avail_m].sum().reset_index(); _pva["期間"] = period_prv
            combined_va = pd.concat([_cva, _pva], ignore_index=True)
        COLOR_MAP = {period_cur: "#4e79a7", period_prv: "#f28e2b"}
        for metric in vendor_metrics:
            if metric not in combined_va.columns:
                continue
            fig = px.bar(
                combined_va[["廠商","期間",metric]].dropna(subset=[metric]),
                x="廠商", y=metric, color="期間", barmode="group",
                title=f"{metric}｜廠商同期比較（{period_cur}" + (f" vs {period_prv}" if has_prv else "") + "）",
                color_discrete_map=COLOR_MAP, text=metric,
            )
            fig.update_traces(texttemplate="%{text:,}", textposition="outside")
            fig.update_layout(height=420, legend_title="期間")
            apply_vlabel(fig, metric)
            st.plotly_chart(fig, use_container_width=True)

    if charts_to_render:
        st.markdown("---")

    # ── 詳細資料
    hdr, tog = st.columns([3, 2])
    hdr.markdown("#### 詳細資料")
    collapse = tog.toggle("📁 摺疊明細（僅顯示小計）", value=False, key=f"{name}_collapse")

    show_cols = [brand_col,"ERP品號","品名","上線量","回廠量",
                 "良品數","再使用率(%)","不良品數","不良率(%)",
                 "過保數","過保率(%)","已使用年限","整體不良率(%)","整體過保率(%)"]
    show_cols = [c for c in show_cols if c in df_cur.columns]
    num_cols  = [c for c in ["上線量","回廠量","良品數","不良品數","過保數"] if c in df_cur.columns]

    sub = df_cur[show_cols].groupby(brand_col, sort=False)[num_cols].sum().reset_index()
    for col_ in ["ERP品號","品名"]:
        if col_ in show_cols: sub[col_] = "── 小計 ──" if col_ == "ERP品號" else ""
    sub["再使用率(%)"]   = (sub["良品數"]   / sub["回廠量"] * 100).fillna(0).round(1)
    sub["不良率(%)"]     = (sub["不良品數"] / sub["回廠量"] * 100).fillna(0).round(1)
    sub["過保率(%)"]     = (sub["過保數"]   / sub["回廠量"] * 100).fillna(0).round(1)
    sub["整體不良率(%)"] = (sub["不良品數"] / sub["上線量"] * 100).fillna(0).round(1)
    sub["整體過保率(%)"] = (sub["過保數"]   / sub["上線量"] * 100).fillna(0).round(1)

    frames = []
    for brand, grp in df_cur[show_cols].groupby(brand_col, sort=False):
        frames.append(grp)
        frames.append(sub[sub[brand_col] == brand])

    total = {c: df_cur[c].sum() for c in num_cols}
    total[brand_col] = "★ 總計"
    if "ERP品號" in show_cols: total["ERP品號"] = ""
    if "品名"    in show_cols: total["品名"]    = ""
    for k in ["再使用率(%)","不良率(%)","過保率(%)","整體不良率(%)","整體過保率(%)"]:
        if k in show_cols:
            n = "良品數" if "再使用" in k else ("不良品數" if "不良" in k else "過保數")
            d = "回廠量" if "過保率" not in k and "不良率(%)" == k or "再使用" in k else ("回廠量" if k in ["不良率(%)","過保率(%)"] else "上線量")
            if n in total and d in total and total[d]:
                total[k] = round(total[n] / total[d] * 100, 1)
    frames.append(pd.DataFrame([total]))

    disp_df   = pd.concat(frames, ignore_index=True)[show_cols]
    is_sub    = disp_df["ERP品號"] == "── 小計 ──" if "ERP品號" in disp_df.columns else pd.Series([False]*len(disp_df))
    is_tot    = disp_df[brand_col] == "★ 總計"
    disp_df   = disp_df.rename(columns={brand_col:"類型","已使用年限":"過保已使用年限(平均)"})

    if collapse:
        view_df  = disp_df[is_sub | is_tot].reset_index(drop=True)
        is_sub_v = view_df["ERP品號"] == "── 小計 ──" if "ERP品號" in view_df.columns else pd.Series([False]*len(view_df))
        is_tot_v = view_df["類型"] == "★ 總計"
        disp_cols = [c for c in view_df.columns if c not in ["ERP品號","品名"]]
        view_df  = view_df[disp_cols]
    else:
        view_df  = disp_df.reset_index(drop=True)
        is_sub_v = is_sub.reset_index(drop=True)
        is_tot_v = is_tot.reset_index(drop=True)

    int_cols = [c for c in num_cols if c in view_df.columns]
    for c in int_cols:
        view_df[c] = pd.to_numeric(view_df[c], errors="coerce").fillna(0).astype(int)

    def highlight(row):
        idx = row.name
        if is_tot_v.iloc[idx]: return ["background-color:#1f4e79;color:white;font-weight:bold"] * len(row)
        if is_sub_v.iloc[idx]: return ["background-color:#d6e4f0;font-weight:bold"] * len(row)
        return [""] * len(row)

    fmt = {c: "{:,}" for c in int_cols}
    fmt.update({k: "{:.1f}" for k in ["再使用率(%)","不良率(%)","過保率(%)","整體不良率(%)","整體過保率(%)","過保已使用年限(平均)"] if k in view_df.columns})

    styled = view_df.style.apply(highlight, axis=1).format(fmt, na_rep="")
    st.dataframe(styled, use_container_width=True, height=max(200, min(60*len(view_df)+38, 600)))

    def to_xlsx(d):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w: d.to_excel(w, index=False)
        return buf.getvalue()

    st.download_button("⬇️ 下載篩選結果 (.xlsx)",
                       data=to_xlsx(df_cur[show_cols]),
                       file_name=f"{name}_{period_cur}_篩選結果.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")

    st.markdown("#### 💡 品管 & 採購建議")
    st.caption("系統依指標自動生成初稿，可直接在文字框內修改後下載使用。")

    rec_tab1, rec_tab2 = st.tabs(["🔍 品管建議", "🛒 採購建議"])
    with rec_tab1:
        auto_q = gen_quality_rec(kpi_c, kpi_p, period_cur, period_prv, name, fault_cols, df_cur)
        eq = st.text_area("品管建議（可自由修改）", value=auto_q, height=300, key=f"{name}_q")
        c1, c2 = st.columns([1, 4])
        c1.download_button("⬇️ 下載 .txt", data=eq.encode("utf-8"),
                           file_name=f"{name}_{period_cur}_品管建議.txt",
                           mime="text/plain", key=f"{name}_dlq")
        if c2.button("🔄 重新生成（清除手動修改）", key=f"{name}_rq"):
            st.rerun()

    with rec_tab2:
        auto_p = gen_purchase_rec(kpi_c, kpi_p, period_cur, period_prv, name, df_cur)
        ep = st.text_area("採購建議（可自由修改）", value=auto_p, height=300, key=f"{name}_p")
        c1, c2 = st.columns([1, 4])
        c1.download_button("⬇️ 下載 .txt", data=ep.encode("utf-8"),
                           file_name=f"{name}_{period_cur}_採購建議.txt",
                           mime="text/plain", key=f"{name}_dlp")
        if c2.button("🔄 重新生成（清除手動修改）", key=f"{name}_rp"):
            st.rerun()


# ── 主 Tabs ───────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["🚗 車機分析", "📷 鏡頭分析"])
with tab1:
    render_tab("車機_彙整總覽", "車機_月趨勢", "車機", CAR_FAULT_COLS)
with tab2:
    render_tab("鏡頭_彙整總覽", "鏡頭_月趨勢", "鏡頭", LENS_FAULT_COLS)

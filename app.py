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
    "上傳設備品質分析 Excel（多期合併版）", type=["xlsx"],
    help="需含：車機_彙整總覽、鏡頭_彙整總覽（含「期間」欄）；選填：車機_月趨勢、鏡頭_月趨勢",
)
if not uploaded_file:
    st.info(
        "請上傳多期合併版分析總表。\n\n"
        "必要工作表：**車機_彙整總覽**、**鏡頭_彙整總覽**（需含 `期間` 欄，如 `2025Q1`、`2026Q1`）\n\n"
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

# ── Y 軸垂直文字（annotation 方式，每字換行）────────────────────────────────
def apply_vlabel(fig, text, margin_l=75):
    stacked = "<br>".join(list(str(text)))
    fig.update_yaxes(title_text="")
    fig.update_layout(margin={"l": margin_l})
    fig.add_annotation(
        text=stacked, showarrow=False,
        xref="paper", yref="paper",
        x=-0.045, y=0.5, xanchor="center", yanchor="middle",
        font=dict(size=12),
    )

# ── 資料前處理 ────────────────────────────────────────────────────────────────
def process(df, fault_cols):
    df = df.copy()
    df.columns = df.columns.str.strip()
    if "期間" not in df.columns:
        df["期間"] = "未知期間"

    good_col  = next((c for c in ["回廠良品數", "良品數"]   if c in df.columns), None)
    bad_col   = next((c for c in ["回廠不良品數", "不良品數"] if c in df.columns), None)
    scrap_col = next((c for c in ["回廠過保數", "過保數"]   if c in df.columns), None)

    if good_col is None:
        df["回廠良品數"] = df[[c for c in ["O /測試正常", "回廠QC", "回廠其他"] if c in df.columns]].sum(axis=1)
        good_col = "回廠良品數"
    if bad_col is None:
        df["回廠不良品數"] = df[[c for c in ["G /評估後退修", "X /已完修", "維修換貨＋換貨條碼"] if c in df.columns]].sum(axis=1)
        bad_col = "回廠不良品數"
    if scrap_col is None:
        df["回廠過保數"] = df[[c for c in ["D /停產報廢", "E /過保報廢", "回廠報廢"] if c in df.columns]].sum(axis=1)
        scrap_col = "回廠過保數"
    if "回廠量" not in df.columns:
        df["回廠量"] = df[[c for c in fault_cols if c in df.columns]].sum(axis=1)

    df["良品數"]  = df[good_col].fillna(0)
    df["不良品數"] = df[bad_col].fillna(0)
    df["過保數"]  = df[scrap_col].fillna(0)

    # 人為數（若原始欄位存在則使用，否則補 0）
    _hum_src = next((c for c in ["人為數", "人為"] if c in df.columns), None)
    df["人為數"] = df[_hum_src].fillna(0) if _hum_src else 0

    df["再使用率(%)"]  = (df["良品數"]  / df["回廠量"] * 100).fillna(0).round(1)
    df["不良率(%)"]   = (df["不良品數"] / df["回廠量"] * 100).fillna(0).round(1)
    df["過保率(%)"]   = (df["過保數"]  / df["回廠量"] * 100).fillna(0).round(1)
    df["整體不良率(%)"] = (df["不良品數"] / df["上線量"] * 100).fillna(0).round(1)
    df["整體過保率(%)"] = (df["過保數"]  / df["上線量"] * 100).fillna(0).round(1)

    # 仍在線數 = 上線量 - 回廠量（供圓環圖使用）
    df["仍在線數"] = (df["上線量"] - df["回廠量"]).clip(lower=0)

    return df

# ── KPI 計算 ──────────────────────────────────────────────────────────────────
def calc_kpi(df):
    if df is None or len(df) == 0:
        return None
    ton  = int(df["上線量"].sum())
    tret = int(df["回廠量"].sum())
    tbad = int(df["不良品數"].sum())
    tgd  = int(df["良品數"].sum())
    tscr = int(df["過保數"].sum())  if "過保數"  in df.columns else 0
    thum = int(df["人為數"].sum())  if "人為數"  in df.columns else 0
    return dict(
        total_on=ton, total_ret=tret,
        total_bad=tbad, total_good=tgd, total_scrap=tscr, total_human=thum,
        reuse_rate   = tgd  / tret * 100 if tret else 0,
        bad_rate     = tbad / tret * 100 if tret else 0,
        scrap_rate   = tscr / tret * 100 if tret else 0,
        ovr_bad_rate = tbad / ton  * 100 if ton  else 0,
        ovr_scr_rate = tscr / ton  * 100 if ton  else 0,
    )

# ── 品管建議生成 ──────────────────────────────────────────────────────────────
def gen_quality_rec(kpi_c, kpi_p, pc, pp, name, fault_cols, df_c):
    pp_label = pp if kpi_p else "（無對比資料）"
    lines = [f"【{name} 品管建議】 {pc} vs {pp_label}\n" + "─"*40]

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

    scr  = kpi_c["ovr_scr_rate"]
    ds   = delta(kpi_c, kpi_p, "ovr_scr_rate")
    ds_s = f"（較去年同期 {'+' if ds and ds>=0 else ''}{ds:.1f}%）" if ds is not None else ""
    if scr >= 3:
        lines.append(f"⚠️ 整體過保率 {scr:.1f}%{ds_s}，建議統計過保品平均使用年限，作為下次汰換計畫依據，並檢視保固條款是否需調整。")
    else:
        lines.append(f"✅ 整體過保率 {scr:.1f}%{ds_s}，目前過保狀況在可控範圍內。")

    lines.append("\n*以上建議為系統自動生成，請依實際情況修改後使用。*")
    return "\n".join(lines)

# ── 採購建議生成 ──────────────────────────────────────────────────────────────
def gen_purchase_rec(kpi_c, kpi_p, pc, pp, name, df_c):
    pp_label  = pp if kpi_p else "（無對比資料）"
    lines     = [f"【{name} 採購建議】 {pc} vs {pp_label}\n" + "─"*40]
    brand_col = next((c for c in ["廠牌型號", "廠牌", "類型"] if c in df_c.columns), df_c.columns[0])

    if "已使用年限" in df_c.columns and "過保率(%)" in df_c.columns:
        hi = df_c[(df_c["過保率(%)"] >= 10) & (df_c["已使用年限"].notna())]
        if not hi.empty:
            items = hi.nlargest(3, "過保率(%)")[brand_col].tolist()
            lines.append(f"🔄 下列類型過保率偏高（≥10%），建議優先列入汰換採購計畫：\n  {'、'.join(map(str, items))}")
        else:
            lines.append("✅ 目前無高過保率品項需緊急汰換。")
    else:
        lines.append("（需有「已使用年限」及「過保率(%)」欄位，方可自動判斷汰換建議）")

    if "整體不良率(%)" in df_c.columns:
        hi2 = df_c[df_c["整體不良率(%)"] >= 5]
        if not hi2.empty:
            items2 = hi2.nlargest(3, "整體不良率(%)")[brand_col].tolist()
            lines.append(f"📋 下列類型整體不良率偏高（≥5%），建議要求廠商提出品質改善計畫，或於下次採購時評估替換供應商：\n  {'、'.join(map(str, items2))}")

    if kpi_p and kpi_p["total_ret"] > 0:
        delta_pct = (kpi_c["total_ret"] - kpi_p["total_ret"]) / kpi_p["total_ret"] * 100
        if delta_pct >= 15:
            lines.append(f"📈 回廠量較去年同期上升 {delta_pct:.1f}%，建議採購合約加入品質保證條款，並提高新批次 AQL 抽樣比例。")
        elif delta_pct <= -10:
            lines.append(f"📉 回廠量較去年同期下降 {abs(delta_pct):.1f}%，品質改善明顯，可維持現行採購策略，適時與供應商確認品質保持計畫。")

    lines.append("\n*以上建議為系統自動生成，請依實際情況修改後使用。*")
    return "\n".join(lines)

# ── 圓環圖：上線量品質占比 ────────────────────────────────────────────────────
def make_quality_donut_online(kpi, title_prefix):
    if kpi is None:
        return None
    online_remaining = max(0, kpi["total_on"] - kpi["total_ret"])
    _cmap = {"仍在線數":"#1f4e79","良品數":"#4472c4","過保數":"#7ba7d3","不良品數":"#a9c4e4","人為數":"#d6e4f0"}
    _all = [("仍在線數", online_remaining), ("良品數", kpi["total_good"]),
            ("過保數", kpi["total_scrap"]), ("不良品數", kpi["total_bad"]),
            ("人為數", kpi["total_human"])]
    _all = [(l, v) for l, v in _all if v > 0]
    labels = [x[0] for x in _all]; values = [x[1] for x in _all]
    marker_colors = [_cmap[l] for l in labels]
    total = kpi["total_on"] or 1
    texts = []
    for lbl, v in zip(labels, values):
        pct = v / total * 100
        if pct >= 10:
            texts.append(f"{lbl}<br>{pct:.1f}%")
        elif pct >= 1:
            texts.append(f"{pct:.1f}%")
        else:
            texts.append("")
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=marker_colors),
        text=texts, textinfo="text",
        textposition="inside",
        insidetextorientation="horizontal",
        textfont=dict(size=11), hoverinfo="label+value+percent",
    ))
    fig.add_annotation(
        text=f"上線總量<br>{kpi['total_on']:,}",
        showarrow=False, font=dict(size=11),
        x=0.5, y=0.5, xanchor="center", yanchor="middle",
    )
    fig.update_layout(
        title=f"{title_prefix} - 上線量品質占比",
        height=380, showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.02,
                    xanchor="center", x=0.5, font=dict(size=10)),
        margin=dict(t=50, b=10, l=10, r=10),
    )
    return fig

# ── 圓環圖：派工回廠品質占比 ──────────────────────────────────────────────────
def make_quality_donut_return(kpi, title_prefix):
    if kpi is None:
        return None
    _cmap = {"良品數":"#1f4e79","過保數":"#4472c4","不良品數":"#7ba7d3","人為數":"#d6e4f0"}
    _all = [("良品數", kpi["total_good"]), ("過保數", kpi["total_scrap"]),
            ("不良品數", kpi["total_bad"]), ("人為數", kpi["total_human"])]
    _all = [(l, v) for l, v in _all if v > 0]
    labels = [x[0] for x in _all]; values = [x[1] for x in _all]
    marker_colors = [_cmap[l] for l in labels]
    total = kpi["total_ret"] or 1
    texts = []
    for lbl, v in zip(labels, values):
        pct = v / total * 100
        if pct >= 10:
            texts.append(f"{lbl}<br>{pct:.1f}%")
        elif pct >= 1:
            texts.append(f"{pct:.1f}%")
        else:
            texts.append("")
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=marker_colors),
        text=texts, textinfo="text",
        textposition="inside",
        insidetextorientation="horizontal",
        textfont=dict(size=11), hoverinfo="label+value+percent",
    ))
    fig.add_annotation(
        text=f"回廠總量<br>{kpi['total_ret']:,}",
        showarrow=False, font=dict(size=11),
        x=0.5, y=0.5, xanchor="center", yanchor="middle",
    )
    fig.update_layout(
        title=f"{title_prefix} - 派工回廠品質占比",
        height=380, showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.02,
                    xanchor="center", x=0.5, font=dict(size=10)),
        margin=dict(t=50, b=10, l=10, r=10),
    )
    return fig

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
        prv_opts   = [p for p in all_periods if p != period_cur]
        period_prv = st.selectbox(
            "對比期間（去年同期）",
            prv_opts if prv_opts else ["（尚無對比資料）"],
            index=0, key=f"{name}_prv",
        )
    has_prv = bool(prv_opts)

    df_cur = df_all[df_all["期間"] == period_cur].copy()
    df_prv = df_all[df_all["期間"] == period_prv].copy() if has_prv else None

    brand_col  = next((c for c in ["廠牌型號", "廠牌", "類型"] if c in df_cur.columns), df_cur.columns[0])
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
                "廠商篩選", sorted(df_cur["廠商"].dropna().astype(str).unique().tolist()),
                key=f"{name}_vendor")
        _v_df = df_cur[df_cur["廠商"].isin(sel_vendors)] if sel_vendors else df_cur
        with fc2:
            sel_brands = st.multiselect(
                "類型篩選", sorted(_v_df[brand_col].dropna().astype(str).unique().tolist()),
                key=f"{name}_brand")
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
                "類型篩選", df_cur[brand_col].dropna().unique(), key=f"{name}_brand")
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
                key=f"{name}_kpi_sel", label_visibility="collapsed")
    if not shown_kpis:
        shown_kpis = DEFAULT_KPIS

    def d_str(cur_v, prv_v, fmt=",.0f", suffix=""):
        if prv_v is None:
            return None
        d = cur_v - prv_v
        return f"{'+' if d >= 0 else ''}{d:{fmt}}{suffix}"

    kpis_all = [
        ("總上線量",    f"{kpi_c['total_on']:,}",        d_str(kpi_c['total_on'],     kpi_p['total_on']     if kpi_p else None)),
        ("期間回廠量",  f"{kpi_c['total_ret']:,}",       d_str(kpi_c['total_ret'],    kpi_p['total_ret']    if kpi_p else None)),
        ("期間再使用率", f"{kpi_c['reuse_rate']:.1f}%",  d_str(kpi_c['reuse_rate'],   kpi_p['reuse_rate']   if kpi_p else None, fmt=".1f", suffix="%")),
        ("期間不良率",  f"{kpi_c['bad_rate']:.1f}%",     d_str(kpi_c['bad_rate'],     kpi_p['bad_rate']     if kpi_p else None, fmt=".1f", suffix="%")),
        ("整體不良率",  f"{kpi_c['ovr_bad_rate']:.1f}%", d_str(kpi_c['ovr_bad_rate'], kpi_p['ovr_bad_rate'] if kpi_p else None, fmt=".1f", suffix="%")),
        ("整體過保率",  f"{kpi_c['ovr_scr_rate']:.1f}%", d_str(kpi_c['ovr_scr_rate'], kpi_p['ovr_scr_rate'] if kpi_p else None, fmt=".1f", suffix="%")),
    ]
    kpis = [(lbl, val, dlt) for lbl, val, dlt in kpis_all if lbl in shown_kpis]
    for i in range(0, len(kpis), 3):
        row_cols = st.columns(3)
        for j, col in enumerate(row_cols):
            if i + j < len(kpis):
                lbl, val, dlt = kpis[i + j]
                col.metric(lbl, val, dlt)

    st.markdown("---")

    # ── 趨勢 & 圖表：統一彈性顯示控制
    df_trend_raw    = get_sheet(trend_sheet_name)
    trend_available = not df_trend_raw.empty
    trend_opt       = ["月趨勢"] if trend_available else []

    base_chart_opts   = ["上線量品質占比", "派工回廠品質占比", "回廠原因分佈", "處置結果分佈"]
    vendor_chart_opts = ["廠商趨勢比較"] if has_vendor else []
    all_display_opts  = trend_opt + base_chart_opts + vendor_chart_opts

    sec_hdr, sec_cfg = st.columns([4, 2])
    sec_hdr.markdown("#### 趨勢與圖表")
    with sec_cfg:
        with st.expander("⚙️ 選擇顯示內容"):
            sel_display = st.multiselect(
                "顯示內容", all_display_opts,
                default=trend_opt + base_chart_opts,
                key=f"{name}_display", label_visibility="collapsed")
    if not sel_display:
        sel_display = trend_opt + base_chart_opts

    # ── 月趨勢（廠商→類型→ERP品號 階層篩選）
    if "月趨勢" in sel_display and trend_available:
        st.markdown("##### 📈 月趨勢")
        df_trend_raw.columns = df_trend_raw.columns.str.strip()
        trend_cols_avail = [c for c in ["回廠量", "不良品數", "良品數", "過保數", "上線量", "已使用年限"] if c in df_trend_raw.columns]
        trend_periods    = [period_cur] + ([period_prv] if has_prv else [])
        dt_base = df_trend_raw[df_trend_raw["期間"].isin(trend_periods)].copy()

        t_brand_col = next((c for c in ["廠牌型號", "廠牌", "類型"] if c in dt_base.columns), None)
        if has_vendor and t_brand_col and "廠商" not in dt_base.columns:
            _join_col = brand_col if brand_col in dt_base.columns else t_brand_col
            _vmap = df_all[["廠商", brand_col]].drop_duplicates().rename(columns={brand_col: _join_col})
            if _join_col in dt_base.columns:
                dt_base = dt_base.merge(_vmap, on=_join_col, how="left")

        t_has_erp = "ERP品號" in dt_base.columns
        if t_has_erp and has_name:
            if "品名" not in dt_base.columns:
                _en_map = (df_all[["ERP品號", "品名"]].dropna(subset=["ERP品號"]).drop_duplicates(subset=["ERP品號"]))
                dt_base = dt_base.merge(_en_map, on="ERP品號", how="left")
            dt_base["_t_label"] = dt_base.apply(
                lambda r: f"{str(r['ERP品號']).strip()} - {str(r['品名']).strip()}"
                          if pd.notna(r["品名"]) and str(r["品名"]).strip() else str(r["ERP品號"]), axis=1)

        t_has_label  = "_t_label" in dt_base.columns
        t_has_vendor = "廠商" in dt_base.columns

        sel_metric = st.selectbox("趨勢指標", trend_cols_avail, key=f"{name}_tm")
        dt = dt_base.copy()
        if sel_vendors and t_has_vendor:
            dt = dt[dt["廠商"].isin(sel_vendors)]
        if sel_brands and t_brand_col:
            dt = dt[dt[t_brand_col].isin(sel_brands)]
        if sel_erp and t_has_erp:
            _filter_col = "_t_label" if t_has_label else "ERP品號"
            dt = dt[dt[_filter_col].isin(sel_erp)]

        if "年月" in dt.columns and sel_metric in dt.columns:
            dt_agg = dt.groupby(["期間", "年月"])[sel_metric].sum().reset_index()
            if dt_agg.empty:
                st.info("目前篩選條件在月趨勢工作表中無對應資料，請確認 ERP品號 是否有月趨勢明細。")
            else:
                dt_agg["月份"] = dt_agg["年月"].astype(str).str[-2:] + "月"
                fig_t = px.line(
                    dt_agg, x="月份", y=sel_metric, color="期間", markers=True,
                    title=f"{name}｜{sel_metric} 月趨勢",
                    color_discrete_map={period_cur: "#4e79a7", period_prv: "#f28e2b"},
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
                "例：`2026Q1` | `2026-01` | `16-1車機` | 120 | 5 | 2 | 3 | 0"
            )

    # ── 圖表（由 sel_display 控制）
    charts_to_render = [c for c in sel_display if c != "月趨勢"]
    if charts_to_render:
        st.markdown("##### 📊 圖表分析")

        # ── 上線量品質占比 & 派工回廠品質占比
        show_online = "上線量品質占比" in charts_to_render
        show_return = "派工回廠品質占比" in charts_to_render
        if show_online or show_return:
            st.markdown("###### 品質占比概覽")
            if show_online and show_return:
                c_tbl1, c_pie1, c_tbl2, c_pie2 = st.columns([1.2, 2, 1.2, 2])
            elif show_online:
                c_tbl1, c_pie1 = st.columns([1, 2])
                c_tbl2 = c_pie2 = None
            else:
                c_tbl2, c_pie2 = st.columns([1, 2])
                c_tbl1 = c_pie1 = None

            if show_online and c_tbl1 is not None:
                online_remaining = max(0, kpi_c["total_on"] - kpi_c["total_ret"])
                _rows_on = [("良品數", kpi_c["total_good"]), ("不良品數", kpi_c["total_bad"]),
                            ("過保數", kpi_c["total_scrap"]), ("仍在線數", online_remaining),
                            ("指定上線總量", kpi_c["total_on"])]
                if kpi_c["total_human"] > 0:
                    _rows_on.insert(3, ("人為數", kpi_c["total_human"]))
                _html_on = (
                    "<table style='font-size:13px;border-collapse:collapse'>"
                    "<tr><th style='text-align:left;padding:3px 10px 3px 4px;border-bottom:1px solid #ddd'>上線車機</th>"
                    "<th style='text-align:right;padding:3px 4px 3px 10px;border-bottom:1px solid #ddd'>總計</th></tr>"
                    + "".join(
                        f"<tr><td style='text-align:left;padding:2px 10px 2px 4px'>{r}</td>"
                        f"<td style='text-align:right;padding:2px 4px 2px 10px'>{v:,}</td></tr>"
                        for r, v in _rows_on
                    ) + "</table>"
                )
                with c_tbl1:
                    st.markdown(_html_on, unsafe_allow_html=True)
                    on_pct  = kpi_c["total_on"]
                    scr_pct = kpi_c["total_scrap"] / on_pct * 100 if on_pct else 0
                    bad_pct = kpi_c["total_bad"]   / on_pct * 100 if on_pct else 0
                    st.caption(f"▲ 可以得知{name}整體上線量占比　過保 僅占{scr_pct:.0f}%　不良品僅占{bad_pct:.0f}%")
                with c_pie1:
                    fig_online = make_quality_donut_online(kpi_c, name)
                    if fig_online:
                        st.plotly_chart(fig_online, use_container_width=True)

            if show_return and c_tbl2 is not None:
                _rows_ret = [("良品數", kpi_c["total_good"]), ("不良品數", kpi_c["total_bad"]),
                             ("過保數", kpi_c["total_scrap"]), ("回廠總量", kpi_c["total_ret"])]
                # 人為數 > 0 才加入
                if kpi_c["total_human"] > 0:
                    _rows_ret.insert(3, ("人為數", kpi_c["total_human"]))
                _html_ret = (
                    "<table style='font-size:13px;border-collapse:collapse'>"
                    "<tr><th style='text-align:left;padding:3px 10px 3px 4px;border-bottom:1px solid #ddd'>回廠車機</th>"
                    "<th style='text-align:right;padding:3px 4px 3px 10px;border-bottom:1px solid #ddd'>總計</th></tr>"
                    + "".join(
                        f"<tr><td style='text-align:left;padding:2px 10px 2px 4px'>{r}</td>"
                        f"<td style='text-align:right;padding:2px 4px 2px 10px'>{v:,}</td></tr>"
                        for r, v in _rows_ret
                    )
                    + "</table>"
                )
                with c_tbl2:
                    st.markdown(_html_ret, unsafe_allow_html=True)
                    ret = kpi_c["total_ret"]
                    gd_pct = kpi_c["total_good"] / ret * 100 if ret else 0
                    st.caption(f"▲ 可以得知{name}派工的回廠比，大多約良品，實際占了{gd_pct:.0f}%")
                with c_pie2:
                    fig_return = make_quality_donut_return(kpi_c, name)
                    if fig_return:
                        st.plotly_chart(fig_return, use_container_width=True)

        # ── 回廠原因分佈
        if "回廠原因分佈" in charts_to_render:
            st.markdown("###### 回廠原因分佈")
            left, right = st.columns(2)
            for col_w, df_w, label in [(left, df_cur, period_cur), (right, df_prv, period_prv if has_prv else None)]:
                if df_w is None or len(df_w) == 0:
                    if label:
                        col_w.info(f"{label}：無資料")
                    continue
                avail  = [c for c in fault_cols if c in df_w.columns]
                totals = df_w[avail].sum()
                totals = totals[totals > 0]
                if not totals.empty:
                    fig = px.pie(values=totals.values, names=totals.index, hole=0.45,
                                 title=f"{name}｜回廠原因（{label}）")
                    fig.update_traces(textinfo="percent+label", textposition="outside")
                    fig.update_layout(height=400)
                    col_w.plotly_chart(fig, use_container_width=True)

        # ── 處置結果分佈
        if "處置結果分佈" in charts_to_render:
            st.markdown("###### 處置結果分佈")
            left, right = st.columns(2)
            for col_w, kpi_w, label in [(left, kpi_c, period_cur), (right, kpi_p, period_prv if has_prv else None)]:
                if kpi_w is None:
                    if label:
                        col_w.info(f"{label}：無資料")
                    continue
                disp = {"良品（再使用）": kpi_w["total_good"], "不良品（維修/換貨）": kpi_w["total_bad"], "過保/報廢": kpi_w["total_scrap"]}
                disp = {k: v for k, v in disp.items() if v > 0}
                if disp:
                    _cm = {"良品（再使用）": "#4472c4", "不良品（維修/換貨）": "#e74c3c", "過保/報廢": "#95a5a6"}
                    fig = px.pie(values=list(disp.values()), names=list(disp.keys()), hole=0.45,
                                 color_discrete_map=_cm, title=f"{name}｜處置結果（{label}）")
                    fig.update_traces(textinfo="percent+label", textposition="outside")
                    fig.update_layout(height=400)
                    col_w.plotly_chart(fig, use_container_width=True)

        # ── 廠商趨勢比較（多廠商跨期折線圖）
        if "廠商趨勢比較" in charts_to_render and has_vendor:
            st.markdown("###### 廠商指標跨期趨勢比較")
            all_vendor_list = sorted(df_all["廠商"].dropna().astype(str).unique().tolist())
            vt_c1, vt_c2 = st.columns(2)
            with vt_c1:
                trend_vendors = st.multiselect(
                    "選擇比較廠商（最多 5 間）", all_vendor_list,
                    max_selections=5, key=f"{name}_trend_vendors")
            with vt_c2:
                _tm_opts = [c for c in ["不良品數", "良品數", "過保數", "回廠量", "整體不良率(%)", "不良率(%)"] if c in df_all.columns]
                trend_metric = st.selectbox("趨勢指標", _tm_opts, key=f"{name}_trend_metric")

            if trend_vendors:
                _ta = df_all[df_all["廠商"].isin(trend_vendors)].copy()
                if sel_brands:
                    _ta = _ta[_ta[brand_col].isin(sel_brands)]
                _ta_agg = (_ta.groupby(["廠商", "期間"])[trend_metric]
                           .sum().reset_index().sort_values("期間"))

                fig_vt = px.line(
                    _ta_agg, x="期間", y=trend_metric, color="廠商",
                    markers=True,
                    title=f"{name}｜廠商 {trend_metric} 跨期趨勢",
                )
                fig_vt.update_traces(line=dict(width=2.5), marker=dict(size=9))

                # 各廠商趨勢箭頭
                for vendor in trend_vendors:
                    vdata = _ta_agg[_ta_agg["廠商"] == vendor].sort_values("期間")
                    if len(vdata) >= 2:
                        last_val = float(vdata[trend_metric].iloc[-1])
                        prev_val = float(vdata[trend_metric].iloc[-2])
                        slope    = last_val - prev_val
                        arrow    = "▲" if slope > 0 else "▼"
                        color    = "red" if slope > 0 else "green"
                        fig_vt.add_annotation(
                            x=vdata["期間"].iloc[-1], y=last_val,
                            text=f"&nbsp;{arrow}",
                            showarrow=False, font=dict(size=18, color=color),
                            xanchor="left", yanchor="middle",
                        )

                fig_vt.update_layout(height=420, legend_title="廠商")
                apply_vlabel(fig_vt, trend_metric)
                st.plotly_chart(fig_vt, use_container_width=True)
            else:
                st.info("請選擇至少一間廠商以顯示趨勢圖。")

        st.markdown("---")

    # ── 鏡頭專屬：室內/室外壽命分析
    if name == "鏡頭":
        st.markdown("#### 🏠/🌿 室內/室外鏡頭壽命分析")
        has_env = "環境類型" in df_cur.columns

        if has_env:
            metric_opts = [c for c in ["已使用年限", "整體不良率(%)", "過保率(%)", "回廠量"] if c in df_cur.columns]
            if not metric_opts:
                st.info("找不到可分析的指標欄位（已使用年限、整體不良率等）。")
            else:
                ev1, _ = st.columns([2, 4])
                with ev1:
                    env_metric = st.selectbox("分析指標", metric_opts, key="lens_env_metric")

                # 箱型圖：壽命/指標分布
                fig_box = px.box(
                    df_cur, x="環境類型", y=env_metric, color="環境類型",
                    title=f"鏡頭｜{env_metric} 分布（室內 vs 室外）",
                    color_discrete_map={"室內": "#4472c4", "室外": "#ed7d31"},
                    points="all",
                )
                fig_box.update_layout(height=400, showlegend=False)
                apply_vlabel(fig_box, env_metric)
                st.plotly_chart(fig_box, use_container_width=True)

                # 使用年限分層長條圖
                if "已使用年限" in df_cur.columns and "回廠量" in df_cur.columns:
                    _env_df = df_cur.copy()
                    _env_df["年限區間"] = pd.cut(
                        _env_df["已使用年限"],
                        bins=[0, 1, 3, 5, float("inf")],
                        labels=["0–1年", "1–3年", "3–5年", "5年以上"],
                        right=False,
                    )
                    _env_agg = _env_df.groupby(["年限區間", "環境類型"])["回廠量"].sum().reset_index()
                    if not _env_agg.empty:
                        fig_bar_env = px.bar(
                            _env_agg, x="年限區間", y="回廠量", color="環境類型",
                            barmode="group",
                            title="鏡頭｜各年限區間回廠量（室內 vs 室外）",
                            color_discrete_map={"室內": "#4472c4", "室外": "#ed7d31"},
                        )
                        fig_bar_env.update_layout(height=400)
                        apply_vlabel(fig_bar_env, "回廠量")
                        st.plotly_chart(fig_bar_env, use_container_width=True)
        else:
            with st.expander("📋 資料格式說明（補齊後自動啟用分析）", expanded=True):
                st.info(
                    "此分析功能需要在 Excel **鏡頭_彙整總覽** 工作表中加入以下欄位：\n\n"
                    "| 欄位名稱 | 說明 | 建議值域 |\n"
                    "|---------|------|----------|\n"
                    "| `環境類型` | 鏡頭安裝環境 | 室內 / 室外 |\n"
                    "| `安裝日期` | 用於自動計算已使用年限 | YYYY-MM-DD |\n\n"
                    "**欄位補齊後將自動顯示：**\n"
                    "- 室內 vs 室外 各指標分布箱型圖\n"
                    "- 使用年限分層（0–1、1–3、3–5、5年以上）回廠量比較\n"
                    "- 各環境類型整體不良率橫向比對\n\n"
                    "若已有「已使用年限」欄位（數值，單位：年），可省略安裝日期。"
                )

        st.markdown("---")

    # ── 詳細資料
    hdr, tog = st.columns([3, 2])
    hdr.markdown("#### 詳細資料")
    collapse = tog.toggle("📁 摺疊明細（僅顯示小計）", value=False, key=f"{name}_collapse")

    show_cols = [brand_col, "ERP品號", "品名", "上線量", "回廠量",
                 "良品數", "再使用率(%)", "不良品數", "不良率(%)",
                 "過保數", "過保率(%)", "已使用年限", "整體不良率(%)", "整體過保率(%)"]
    show_cols = [c for c in show_cols if c in df_cur.columns]
    num_cols  = [c for c in ["上線量", "回廠量", "良品數", "不良品數", "過保數"] if c in df_cur.columns]

    sub = df_cur[show_cols].groupby(brand_col, sort=False)[num_cols].sum().reset_index()
    for col_ in ["ERP品號", "品名"]:
        if col_ in show_cols:
            sub[col_] = "── 小計 ──" if col_ == "ERP品號" else ""
    sub["再使用率(%)"]   = (sub["良品數"]  / sub["回廠量"] * 100).fillna(0).round(1)
    sub["不良率(%)"]    = (sub["不良品數"] / sub["回廠量"] * 100).fillna(0).round(1)
    sub["過保率(%)"]    = (sub["過保數"]  / sub["回廠量"] * 100).fillna(0).round(1)
    sub["整體不良率(%)"] = (sub["不良品數"] / sub["上線量"] * 100).fillna(0).round(1)
    sub["整體過保率(%)"] = (sub["過保數"]  / sub["上線量"] * 100).fillna(0).round(1)

    frames = []
    for brand, grp in df_cur[show_cols].groupby(brand_col, sort=False):
        frames.append(grp)
        frames.append(sub[sub[brand_col] == brand])

    total = {c: df_cur[c].sum() for c in num_cols}
    total[brand_col] = "★ 總計"
    if "ERP品號" in show_cols: total["ERP品號"] = ""
    if "品名"    in show_cols: total["品名"]    = ""
    for k in ["再使用率(%)", "不良率(%)", "過保率(%)", "整體不良率(%)", "整體過保率(%)"]:
        if k in show_cols:
            n = "良品數" if "再使用" in k else ("不良品數" if "不良" in k else "過保數")
            d = "回廠量" if k in ["再使用率(%)", "不良率(%)", "過保率(%)"] else "上線量"
            if n in total and d in total and total[d]:
                total[k] = round(total[n] / total[d] * 100, 1)
    frames.append(pd.DataFrame([total]))

    disp_df = pd.concat(frames, ignore_index=True)[show_cols]
    is_sub  = disp_df["ERP品號"] == "── 小計 ──" if "ERP品號" in disp_df.columns else pd.Series([False] * len(disp_df))
    is_tot  = disp_df[brand_col] == "★ 總計"
    disp_df = disp_df.rename(columns={brand_col: "類型", "已使用年限": "過保已使用年限(平均)"})

    if collapse:
        view_df  = disp_df[is_sub | is_tot].reset_index(drop=True)
        is_sub_v = view_df["ERP品號"] == "── 小計 ──" if "ERP品號" in view_df.columns else pd.Series([False] * len(view_df))
        is_tot_v = view_df["類型"] == "★ 總計"
        disp_cols = [c for c in view_df.columns if c not in ["ERP品號", "品名"]]
        view_df   = view_df[disp_cols]
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
    fmt.update({k: "{:.1f}" for k in ["再使用率(%)", "不良率(%)", "過保率(%)", "整體不良率(%)", "整體過保率(%)", "過保已使用年限(平均)"] if k in view_df.columns})
    styled = view_df.style.apply(highlight, axis=1).format(fmt, na_rep="")
    st.dataframe(styled, use_container_width=True, height=max(200, min(60 * len(view_df) + 38, 600)))

    def to_xlsx(d):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            d.to_excel(w, index=False)
        return buf.getvalue()

    st.download_button(
        "⬇️ 下載篩選結果 (.xlsx)",
        data=to_xlsx(df_cur[show_cols]),
        file_name=f"{name}_{period_cur}_篩選結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("---")

    st.markdown("#### 💡 品管 & 採購建議")
    st.caption("系統依指標自動生成初稿，可直接在文字框內修改後下載使用。")
    rec_tab1, rec_tab2 = st.tabs(["🔍 品管建議", "🛒 採購建議"])
    with rec_tab1:
        auto_q = gen_quality_rec(kpi_c, kpi_p, period_cur, period_prv, name, fault_cols, df_cur)
        eq = st.text_area("品管建議（可自由修改）", value=auto_q, height=300, key=f"{name}_q")
        c1, c2 = st.columns([1, 4])
        c1.download_button("⬇️ 下載 .txt", data=eq.encode("utf-8"),
                           file_name=f"{name}_{period_cur}_品管建議.txt", mime="text/plain", key=f"{name}_dlq")
        if c2.button("🔄 重新生成（清除手動修改）", key=f"{name}_rq"):
            st.rerun()
    with rec_tab2:
        auto_p = gen_purchase_rec(kpi_c, kpi_p, period_cur, period_prv, name, df_cur)
        ep = st.text_area("採購建議（可自由修改）", value=auto_p, height=300, key=f"{name}_p")
        c1, c2 = st.columns([1, 4])
        c1.download_button("⬇️ 下載 .txt", data=ep.encode("utf-8"),
                           file_name=f"{name}_{period_cur}_採購建議.txt", mime="text/plain", key=f"{name}_dlp")
        if c2.button("🔄 重新生成（清除手動修改）", key=f"{name}_rp"):
            st.rerun()

# ── 主 Tabs ───────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["🚗 車機分析", "📷 鏡頭分析"])
with tab1:
    render_tab("車機_彙整總覽", "車機_月趨勢", "車機", CAR_FAULT_COLS)
with tab2:
    render_tab("鏡頭_彙整總覽", "鏡頭_月趨勢", "鏡頭", LENS_FAULT_COLS)

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

uploaded_file = st.file_uploader("上傳設備品質分析 Excel", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.read()

    try:
        df_car  = pd.read_excel(BytesIO(file_bytes), sheet_name="車機_彙整總覽")
        df_lens = pd.read_excel(BytesIO(file_bytes), sheet_name="鏡頭_彙整總覽")
    except Exception as e:
        st.error(f"讀取工作表失敗：{e}")
        st.stop()

    CAR_FAULT_COLS  = ["AB點", "失聯", "定位不良", "訊號異常", "其他"]
    LENS_FAULT_COLS = ["黑畫面", "進水/模糊", "水波紋", "時有時無", "其他"]

    def process(df, fault_cols):
        df = df.copy()
        df.columns = df.columns.str.strip()

        good_col  = next((c for c in ["回廠良品數",  "良品數"]  if c in df.columns), None)
        bad_col   = next((c for c in ["回廠不良品數", "不良品數"] if c in df.columns), None)
        scrap_col = next((c for c in ["回廠過保數",  "過保數"]  if c in df.columns), None)

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

        df["良品數"]   = df[good_col].fillna(0)
        df["不良品數"] = df[bad_col].fillna(0)
        df["過保數"]   = df[scrap_col].fillna(0)

        df["再使用率(%)"]   = (df["良品數"]   / df["回廠量"] * 100).fillna(0).round(1)
        df["不良率(%)"]     = (df["不良品數"] / df["回廠量"] * 100).fillna(0).round(1)
        df["過保率(%)"]     = (df["過保數"]   / df["回廠量"] * 100).fillna(0).round(1)
        df["整體不良率(%)"] = (df["不良品數"] / df["上線量"] * 100).fillna(0).round(1)
        df["整體過保率(%)"] = (df["過保數"]   / df["上線量"] * 100).fillna(0).round(1)
        return df

    df_car  = process(df_car,  CAR_FAULT_COLS)
    df_lens = process(df_lens, LENS_FAULT_COLS)

    # ── KPI 卡片 ────────────────────────────────────────────────────────────
    def kpi_row(df, show_flags):
        total_on    = int(df["上線量"].sum())
        total_ret   = int(df["回廠量"].sum())
        total_bad   = int(df["不良品數"].sum())
        total_good  = int(df["良品數"].sum())
        total_scrap = int(df["過保數"].sum()) if "過保數" in df.columns else 0

        reuse_rate         = total_good  / total_ret * 100 if total_ret else 0
        period_bad_rate    = total_bad   / total_ret * 100 if total_ret else 0
        period_scrap_rate  = total_scrap / total_ret * 100 if total_ret else 0
        overall_bad_rate   = total_bad   / total_on  * 100 if total_on  else 0
        overall_scrap_rate = total_scrap / total_on  * 100 if total_on  else 0

        metrics = []
        if show_flags.get("總上線量",    True): metrics.append(("總上線量",    f"{total_on:,}"))
        if show_flags.get("期間回廠量",  True): metrics.append(("期間回廠量",  f"{total_ret:,}"))
        if show_flags.get("期間再使用率",True): metrics.append(("期間再使用率",f"{reuse_rate:.1f}%"))
        if show_flags.get("期間不良率",  True): metrics.append(("期間不良率",  f"{period_bad_rate:.1f}%"))
        if show_flags.get("期間過保率",  True): metrics.append(("期間過保率",  f"{period_scrap_rate:.1f}%"))
        if show_flags.get("整體不良率",  True): metrics.append(("整體不良率",  f"{overall_bad_rate:.1f}%"))
        if show_flags.get("整體過保率",  True): metrics.append(("整體過保率",  f"{overall_scrap_rate:.1f}%"))

        if metrics:
            cols = st.columns(len(metrics))
            for col, (label, value) in zip(cols, metrics):
                col.metric(label, value)

    # ── 主繪圖函數 ──────────────────────────────────────────────────────────
    def render_tab(df, name, fault_cols):
        brand_col = next(
            (c for c in ["廠牌型號", "廠牌", "類型"] if c in df.columns),
            df.columns[0]
        )

        has_erp  = "ERP品號" in df.columns
        has_name = "品名"    in df.columns

        if has_erp:
            if has_name:
                def make_label(row):
                    pn    = str(row["ERP品號"]).strip() if pd.notna(row["ERP品號"]) else ""
                    name_ = str(row["品名"]).strip()    if pd.notna(row["品名"])    else ""
                    return f"{pn} - {name_}" if name_ else pn
                df = df.copy()
                df["_erp_label"] = df.apply(make_label, axis=1)
                erp_label_col = "_erp_label"
            else:
                erp_label_col = "ERP品號"
        else:
            erp_label_col = None

        # ── 顯示設定 ────────────────────────────────────────────────────────
        with st.expander("⚙️ 顯示設定", expanded=False):
            cfg1, cfg2, cfg3 = st.columns(3)

            with cfg1:
                st.markdown("**🎨 長條圖配色**")
                color_online     = st.color_picker("上線量顏色",        value="#4e79a7", key=f"{name}_c_online")
                color_return     = st.color_picker("回廠量顏色",        value="#f28e2b", key=f"{name}_c_return")
                st.markdown("**不良率漸層**")
                color_bad_low    = st.color_picker("低端色（好）",      value="#00b050", key=f"{name}_c_bad_lo")
                color_bad_high   = st.color_picker("高端色（差）",      value="#ff0000", key=f"{name}_c_bad_hi")
                st.markdown("**再使用率漸層**")
                color_reuse_low  = st.color_picker("低端色（差）",      value="#ff0000", key=f"{name}_c_reu_lo")
                color_reuse_high = st.color_picker("高端色（好）",      value="#00b050", key=f"{name}_c_reu_hi")
                st.markdown("---")
                st.markdown("**🔢 數值顯示模式**")
                label_mode = st.radio(
                    "長條圖數值顯示方式",
                    options=[
                        "A｜外部顯示（預設）",
                        "B｜自動判斷位置",
                        "C｜小字體外部顯示",
                        "D｜僅 Hover 顯示",
                    ],
                    index=0,
                    key=f"{name}_label_mode",
                    help="A：固定外側｜B：自動判斷｜C：字體縮小｜D：僅 Hover",
                )

            with cfg3:
                st.markdown("**🥧 圓環圖配色**")
                st.markdown("**處置結果 / 上線量構成**")
                _fault_defaults = ["#636efa", "#ef553b", "#00cc96", "#ab63fa", "#ffa15a"]
                pie_c_good  = st.color_picker("良品（再使用）",      value="#4472c4", key=f"{name}_pc_good")
                pie_c_bad   = st.color_picker("不良品（維修/換貨）", value="#e74c3c", key=f"{name}_pc_bad")
                pie_c_scrap = st.color_picker("過保/報廢",           value="#95a5a6", key=f"{name}_pc_scr")
                pie_c_still = st.color_picker("仍在線上數",          value="#4e79a7", key=f"{name}_pc_still")
                st.markdown("**回廠原因分佈**")
                pie_fault_colors = {
                    fc: st.color_picker(fc, value=_fault_defaults[i % len(_fault_defaults)],
                                        key=f"{name}_pc_fault_{i}")
                    for i, fc in enumerate(fault_cols)
                }

            with cfg2:
                st.markdown("**📊 整體指標顯示**")
                show_kpi = {
                    "總上線量":    st.checkbox("總上線量",    value=True, key=f"{name}_kpi_on"),
                    "期間回廠量":  st.checkbox("期間回廠量",  value=True, key=f"{name}_kpi_ret"),
                    "期間再使用率":st.checkbox("期間再使用率",value=True, key=f"{name}_kpi_reuse"),
                    "期間不良率":  st.checkbox("期間不良率",  value=True, key=f"{name}_kpi_pbad"),
                    "期間過保率":  st.checkbox("期間過保率",  value=True, key=f"{name}_kpi_pscr"),
                    "整體不良率":  st.checkbox("整體不良率",  value=True, key=f"{name}_kpi_bad"),
                    "整體過保率":  st.checkbox("整體過保率",  value=True, key=f"{name}_kpi_scr"),
                }
                st.markdown("---")
                st.markdown("**📈 圖表顯示設定**")
                show_bar_bad   = st.checkbox("各類型不良率長條圖",    value=False, key=f"{name}_ch_bad")
                show_pie_fault = st.checkbox("回廠原因分佈圓餅圖",    value=True,  key=f"{name}_ch_fault")
                show_bar_reuse = st.checkbox("各類型再使用率長條圖",  value=False, key=f"{name}_ch_reuse")
                show_pie_disp  = st.checkbox("處置結果分佈圓餅圖",    value=True,  key=f"{name}_ch_disp")
                show_bar_onret = st.checkbox("上線量 vs 回廠量長條圖",value=False, key=f"{name}_ch_onret")
                show_pie_comp  = st.checkbox("上線量構成圓餅圖",      value=True,  key=f"{name}_ch_comp")

        # ── 篩選器 ──────────────────────────────────────────────────────────
        f1, f2 = st.columns(2)
        with f1:
            sel_brands = st.multiselect(
                "類型篩選", df[brand_col].dropna().unique(), key=f"{name}_brand"
            )
        brand_filtered = df[df[brand_col].isin(sel_brands)] if sel_brands else df
        with f2:
            if erp_label_col:
                erp_options    = brand_filtered[erp_label_col].dropna().unique().tolist()
                sel_erp_labels = st.multiselect("ERP品號篩選", erp_options, key=f"{name}_erp")
            else:
                sel_erp_labels = []
                st.multiselect("ERP品號篩選", [], key=f"{name}_erp")

        filtered = brand_filtered.copy()
        if sel_erp_labels and erp_label_col:
            filtered = filtered[filtered[erp_label_col].isin(sel_erp_labels)]

        # ── KPI ─────────────────────────────────────────────────────────────
        st.markdown("#### 整體指標")
        kpi_row(filtered, show_kpi)
        st.markdown("---")

        bad_scale   = [[0, color_bad_low],   [1, color_bad_high]]
        reuse_scale = [[0, color_reuse_low],  [1, color_reuse_high]]

        if label_mode.startswith("A"):
            txt_pos, txt_size, show_text = "outside", 12, True
        elif label_mode.startswith("B"):
            txt_pos, txt_size, show_text = "auto",    12, True
        elif label_mode.startswith("C"):
            txt_pos, txt_size, show_text = "outside",  9, True
        else:
            txt_pos, txt_size, show_text = "outside", 12, False

        # ── 圖表渲染函數（接受 container 參數）──────────────────────────────
        def make_bar_bad(ct):
            fig = px.bar(
                filtered.sort_values("整體不良率(%)", ascending=False),
                x=brand_col, y="整體不良率(%)",
                color="整體不良率(%)",
                color_continuous_scale=bad_scale,
                title=f"{name}｜各類型不良率",
                text="整體不良率(%)" if show_text else None,
                labels={brand_col: "類型"},
            )
            if show_text:
                fig.update_traces(texttemplate="%{text:.1f}%", textposition=txt_pos, textfont_size=txt_size)
            fig.update_layout(
                coloraxis_showscale=False, height=420, yaxis_title="",
                annotations=[dict(text="不良率 (%)", x=-0.07, y=1.08,
                                  xref="paper", yref="paper", showarrow=False, font=dict(size=13))],
            )
            ct.plotly_chart(fig, use_container_width=True)

        def make_pie_fault(ct):
            avail = [c for c in fault_cols if c in filtered.columns]
            totals = filtered[avail].sum()
            totals = totals[totals > 0]
            if not totals.empty:
                fig2 = px.pie(values=totals.values, names=totals.index,
                              hole=0.45, color_discrete_map=pie_fault_colors,
                              title=f"{name}｜回廠原因分佈")
                if label_mode.startswith("A"):
                    fig2.update_traces(textinfo="percent+label", textposition="outside", textfont_size=12)
                elif label_mode.startswith("B"):
                    fig2.update_traces(textinfo="percent", textposition="auto",    textfont_size=12)
                elif label_mode.startswith("C"):
                    fig2.update_traces(textinfo="percent", textposition="outside", textfont_size=9)
                else:
                    fig2.update_traces(textinfo="none")
                fig2.update_layout(height=420)
                ct.plotly_chart(fig2, use_container_width=True)
            else:
                ct.info("無回廠原因資料")

        def make_bar_reuse(ct):
            fig3 = px.bar(
                filtered.sort_values("再使用率(%)", ascending=False),
                x=brand_col, y="再使用率(%)",
                color="再使用率(%)",
                color_continuous_scale=reuse_scale,
                title=f"{name}｜各類型再使用率",
                text="再使用率(%)" if show_text else None,
                labels={brand_col: "類型"},
            )
            if show_text:
                fig3.update_traces(texttemplate="%{text:.1f}%", textposition=txt_pos, textfont_size=txt_size)
            fig3.update_layout(
                coloraxis_showscale=False, height=420, yaxis_title="",
                annotations=[dict(text="再使用率 (%)", x=-0.07, y=1.08,
                                  xref="paper", yref="paper", showarrow=False, font=dict(size=13))],
            )
            ct.plotly_chart(fig3, use_container_width=True)

        def make_pie_disp(ct):
            disp_map = {
                "良品數":   "良品（再使用）",
                "不良品數": "不良品（維修/換貨）",
                "過保數":   "過保/報廢",
            }
            vals = {v: int(filtered[k].sum()) for k, v in disp_map.items() if k in filtered.columns}
            vals = {k: v for k, v in vals.items() if v > 0}
            if vals:
                _disp_cmap = {"良品（再使用）": pie_c_good, "不良品（維修/換貨）": pie_c_bad, "過保/報廢": pie_c_scrap}
                fig4 = px.pie(values=list(vals.values()), names=list(vals.keys()),
                              hole=0.45, color_discrete_map=_disp_cmap,
                              title=f"{name}｜處置結果分佈")
                if label_mode.startswith("A"):
                    fig4.update_traces(textinfo="percent+label", textposition="outside", textfont_size=12)
                elif label_mode.startswith("B"):
                    fig4.update_traces(textinfo="percent", textposition="auto",    textfont_size=12)
                elif label_mode.startswith("C"):
                    fig4.update_traces(textinfo="percent", textposition="outside", textfont_size=9)
                else:
                    fig4.update_traces(textinfo="none")
                fig4.update_layout(height=420)
                ct.plotly_chart(fig4, use_container_width=True)
            else:
                ct.info("無處置結果資料")

        def make_bar_onret(ct):
            fig5 = go.Figure()
            kw  = dict(text=filtered["上線量"], texttemplate="%{text:,}", textposition=txt_pos, textfont_size=txt_size) if show_text else {}
            kw2 = dict(text=filtered["回廠量"], texttemplate="%{text:,}", textposition=txt_pos, textfont_size=txt_size) if show_text else {}
            fig5.add_trace(go.Bar(name="上線量", x=filtered[brand_col], y=filtered["上線量"], marker_color=color_online, **kw))
            fig5.add_trace(go.Bar(name="回廠量", x=filtered[brand_col], y=filtered["回廠量"], marker_color=color_return, **kw2))
            fig5.update_layout(barmode="group", height=420, xaxis_title="類型", yaxis_title="數量",
                               title=f"{name}｜上線量 vs 回廠量")
            ct.plotly_chart(fig5, use_container_width=True)

        def make_pie_comp(ct):
            total_on_f   = int(filtered["上線量"].sum())
            total_ret_f  = int(filtered["回廠量"].sum())
            total_good_f = int(filtered["良品數"].sum())
            total_bad_f  = int(filtered["不良品數"].sum())
            total_scr_f  = int(filtered["過保數"].sum()) if "過保數" in filtered.columns else 0
            still_on     = max(total_on_f - total_ret_f, 0)
            comp = {
                "良品（再使用）":      total_good_f,
                "不良品（維修/換貨）": total_bad_f,
                "過保/報廢":           total_scr_f,
                "仍在線上數":          still_on,
            }
            comp = {k: v for k, v in comp.items() if v > 0}
            if comp:
                _comp_cmap = {"良品（再使用）": pie_c_good, "不良品（維修/換貨）": pie_c_bad,
                              "過保/報廢": pie_c_scrap, "仍在線上數": pie_c_still}
                fig6 = px.pie(values=list(comp.values()), names=list(comp.keys()),
                              hole=0.45, color_discrete_map=_comp_cmap,
                              title=f"{name}｜上線量構成")
                if label_mode.startswith("A"):
                    fig6.update_traces(textinfo="percent+label", textposition="outside", textfont_size=12)
                elif label_mode.startswith("B"):
                    fig6.update_traces(textinfo="percent", textposition="auto",    textfont_size=12)
                elif label_mode.startswith("C"):
                    fig6.update_traces(textinfo="percent", textposition="outside", textfont_size=9)
                else:
                    fig6.update_traces(textinfo="none")
                fig6.update_layout(height=420)
                ct.plotly_chart(fig6, use_container_width=True)
            else:
                ct.info("無上線量構成資料")

        # ── 動態兩欄配對佈局 ────────────────────────────────────────────────
        chart_funcs = []
        if show_bar_bad:   chart_funcs.append(make_bar_bad)
        if show_pie_fault: chart_funcs.append(make_pie_fault)
        if show_bar_reuse: chart_funcs.append(make_bar_reuse)
        if show_pie_disp:  chart_funcs.append(make_pie_disp)
        if show_bar_onret: chart_funcs.append(make_bar_onret)
        if show_pie_comp:  chart_funcs.append(make_pie_comp)

        for i in range(0, len(chart_funcs), 2):
            left, right = st.columns(2)
            chart_funcs[i](left)
            if i + 1 < len(chart_funcs):
                chart_funcs[i + 1](right)

        st.markdown("---")

        # ── 詳細資料表（含廠牌小計） ─────────────────────────────────────
        hd_col, tog_col = st.columns([3, 2])
        hd_col.markdown("#### 詳細資料")
        collapse_subtotal = tog_col.toggle(
            "📁 摺疊明細（僅顯示小計）", value=False, key=f"{name}_collapse"
        )

        show_cols = [brand_col, "ERP品號", "品名", "上線量", "回廠量",
                     "良品數", "再使用率(%)", "不良品數", "不良率(%)",
                     "過保數", "過保率(%)", "已使用年限", "整體不良率(%)", "整體過保率(%)"]
        show_cols = [c for c in show_cols if c in filtered.columns]
        num_cols  = ["上線量", "回廠量", "良品數", "不良品數", "過保數"]
        num_cols  = [c for c in num_cols if c in filtered.columns]

        subtotals = (
            filtered[show_cols].groupby(brand_col, sort=False)[num_cols]
            .sum().reset_index()
        )
        subtotals["ERP品號"]       = "── 小計 ──"
        subtotals["品名"]          = ""
        subtotals["再使用率(%)"]   = (subtotals["良品數"]   / subtotals["回廠量"] * 100).fillna(0).round(1)
        subtotals["不良率(%)"]     = (subtotals["不良品數"] / subtotals["回廠量"] * 100).fillna(0).round(1)
        subtotals["過保率(%)"]     = (subtotals["過保數"]   / subtotals["回廠量"] * 100).fillna(0).round(1)
        subtotals["整體不良率(%)"] = (subtotals["不良品數"] / subtotals["上線量"] * 100).fillna(0).round(1)
        subtotals["整體過保率(%)"] = (subtotals["過保數"]   / subtotals["上線量"] * 100).fillna(0).round(1)
        if "已使用年限" in show_cols:
            yr_mean = filtered.groupby(brand_col, sort=False)["已使用年限"].mean().round(1)
            subtotals["已使用年限"] = subtotals[brand_col].map(yr_mean)

        frames = []
        for brand, grp in filtered[show_cols].groupby(brand_col, sort=False):
            frames.append(grp)
            frames.append(subtotals[subtotals[brand_col] == brand])

        total_row = {c: filtered[c].sum() for c in num_cols}
        total_row[brand_col]       = "★ 總計"
        total_row["ERP品號"]       = ""
        total_row["品名"]          = ""
        total_row["再使用率(%)"]   = round(total_row["良品數"]   / total_row["回廠量"] * 100, 1) if total_row.get("回廠量") else 0
        total_row["不良率(%)"]     = round(total_row["不良品數"] / total_row["回廠量"] * 100, 1) if total_row.get("回廠量") else 0
        total_row["過保率(%)"]     = round(total_row["過保數"]   / total_row["回廠量"] * 100, 1) if total_row.get("回廠量") else 0
        total_row["整體不良率(%)"] = round(total_row["不良品數"] / total_row["上線量"] * 100, 1) if total_row.get("上線量") else 0
        total_row["整體過保率(%)"] = round(total_row["過保數"]   / total_row["上線量"] * 100, 1) if total_row.get("上線量") else 0
        if "已使用年限" in show_cols:
            total_row["已使用年限"] = round(filtered["已使用年限"].mean(), 1)
        frames.append(pd.DataFrame([total_row]))

        display_df  = pd.concat(frames, ignore_index=True)[show_cols]
        is_subtotal = display_df["ERP品號"] == "── 小計 ──"
        is_total    = display_df[brand_col]  == "★ 總計"
        display_df  = display_df.rename(columns={brand_col: "類型"})

        if collapse_subtotal:
            view_df     = display_df[is_subtotal | is_total].reset_index(drop=True)
            is_sub_view = view_df["ERP品號"] == "── 小計 ──"
            is_tot_view = view_df["類型"]    == "★ 總計"
        else:
            view_df     = display_df.reset_index(drop=True)
            is_sub_view = is_subtotal.reset_index(drop=True)
            is_tot_view = is_total.reset_index(drop=True)

        int_cols = [c for c in ["上線量", "回廠量", "良品數", "不良品數", "過保數"] if c in view_df.columns]
        for c in int_cols:
            view_df[c] = pd.to_numeric(view_df[c], errors="coerce").fillna(0).astype(int)

        def highlight_rows(row):
            idx = row.name
            if is_tot_view.iloc[idx]:
                return ["background-color: #1f4e79; color: white; font-weight: bold"] * len(row)
            if is_sub_view.iloc[idx]:
                return ["background-color: #d6e4f0; font-weight: bold"] * len(row)
            return [""] * len(row)

        fmt = {c: "{:,}" for c in int_cols}
        fmt.update({
            "再使用率(%)":   "{:.1f}",
            "不良率(%)":     "{:.1f}",
            "過保率(%)":     "{:.1f}",
            "整體不良率(%)": "{:.1f}",
            "整體過保率(%)": "{:.1f}",
        })
        if "已使用年限" in view_df.columns:
            fmt["已使用年限"] = "{:.1f}"

        if collapse_subtotal:
            disp_cols = [c for c in view_df.columns if c not in ["ERP品號", "品名"]]
            view_df   = view_df[disp_cols]
            fmt       = {k: v for k, v in fmt.items() if k in disp_cols}

        table_height = max(200, min(60 * len(view_df) + 38, 600))
        styled = view_df.style.apply(highlight_rows, axis=1).format(fmt, na_rep="")
        st.dataframe(styled, use_container_width=True, height=table_height)

        # ── 下載 ─────────────────────────────────────────────────────────
        def to_excel(d):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                d.to_excel(w, index=False)
            return buf.getvalue()

        st.download_button(
            label="⬇️ 下載篩選結果 (.xlsx)",
            data=to_excel(filtered),
            file_name=f"{name}_篩選結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab1, tab2 = st.tabs(["🚗 車機分析", "📷 鏡頭分析"])
    with tab1:
        render_tab(df_car,  "車機",  CAR_FAULT_COLS)
    with tab2:
        render_tab(df_lens, "鏡頭", LENS_FAULT_COLS)

else:
    st.info("請先上傳 Excel 檔案 - 分析總表（需含「車機、鏡頭 兩個工作表）")

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

# ── 側邊欄：調色盤設定 ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎨 配色設定")
    with st.expander("展開調色盤", expanded=False):
        color_online = st.color_picker("上線量顏色", value="#4e79a7")
        color_return = st.color_picker("回廠量顏色", value="#f28e2b")
        st.markdown("---")
        st.markdown("**不良率漸層**")
        color_bad_low  = st.color_picker("不良率低端色（好）", value="#00b050")
        color_bad_high = st.color_picker("不良率高端色（差）", value="#ff0000")
        st.markdown("**再使用率漸層**")
        color_reuse_low  = st.color_picker("再使用率低端色（差）", value="#ff0000")
        color_reuse_high = st.color_picker("再使用率高端色（好）", value="#00b050")

    st.markdown("## 🔢 數值顯示模式")
    label_mode = st.radio(
        "長條圖數值顯示方式",
        options=[
            "A｜外部顯示（預設）",
            "B｜自動判斷位置",
            "C｜小字體外部顯示",
            "D｜僅 Hover 顯示",
        ],
        index=0,
        help=(
            "A：數值固定顯示於長條外側（原始設定，類型多時可能重疊）\n\n"
            "B：Plotly 自動判斷放內或外，較不易碰撞\n\n"
            "C：字體縮小至 10px，外部顯示，密集時較整齊\n\n"
            "D：長條上不顯示數字，移到長條上才出現數值"
        ),
    )

uploaded_file = st.file_uploader("上傳設備品質分析 Excel", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.read()

    try:
        df_car  = pd.read_excel(BytesIO(file_bytes), sheet_name="車機_彙整總覽")
        df_lens = pd.read_excel(BytesIO(file_bytes), sheet_name="鏡頭_彙整總覽")
    except Exception as e:
        st.error(f"讀取工作表失敗：{e}")
        st.stop()

    # ── 欄位定義（依實際 Excel 欄位名稱） ────────────────────────────────
    CAR_FAULT_COLS  = ["AB點", "失聯", "定位不良", "訊號異常", "其他"]
    LENS_FAULT_COLS = ["黑畫面", "進水/模糊", "水波紋", "時有時無", "其他"]

    def process(df, fault_cols):
        df = df.copy()
        df.columns = df.columns.str.strip()

        # 直接使用 Excel 已有的彙整欄位
        good_col  = next((c for c in ["回廠良品數",  "良品數"]  if c in df.columns), None)
        bad_col   = next((c for c in ["回廠不良品數", "不良品數"] if c in df.columns), None)
        scrap_col = next((c for c in ["回廠過保數",  "過保數"]  if c in df.columns), None)

        # 若 Excel 沒有彙整欄，則從明細欄加總
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

        # 統一用短名稱方便後續引用
        df["良品數"]   = df[good_col].fillna(0)
        df["不良品數"] = df[bad_col].fillna(0)
        df["過保數"]   = df[scrap_col].fillna(0)

        df["不良率(%)"]   = (df["不良品數"] / df["上線量"]  * 100).fillna(0).round(1)
        df["再使用率(%)"] = (df["良品數"]   / df["回廠量"]  * 100).fillna(0).round(1)
        df["過保率(%)"]   = (df["過保數"]   / df["上線量"]  * 100).fillna(0).round(1)
        return df

    df_car  = process(df_car,  CAR_FAULT_COLS)
    df_lens = process(df_lens, LENS_FAULT_COLS)

    # ── KPI 卡片 ────────────────────────────────────────────────────────────
    def kpi_row(df):
        total_on   = int(df["上線量"].sum())
        total_ret  = int(df["回廠量"].sum())
        total_bad  = int(df["不良品數"].sum())
        total_good = int(df["良品數"].sum())
        bad_rate   = total_bad  / total_on  * 100 if total_on  else 0
        reuse_rate = total_good / total_ret * 100 if total_ret else 0
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("總上線量",     f"{total_on:,}")
        c2.metric("總回廠量",     f"{total_ret:,}")
        c3.metric("不良品數",     f"{total_bad:,}")
        c4.metric("整體不良率",   f"{bad_rate:.1f}%")
        c5.metric("整體再使用率", f"{reuse_rate:.1f}%")

    # ── 主繪圖函數 ──────────────────────────────────────────────────────────
    def render_tab(df, name, fault_cols):
        # 動態找品牌欄（廠牌型號 → 廠牌 → 類型 → 第一欄）
        brand_col = next(
            (c for c in ["廠牌型號", "廠牌", "類型"] if c in df.columns),
            df.columns[0]
        )

        # ── 建立 ERP品號+品名 複合顯示標籤 ──────────────────────────────
        has_erp  = "ERP品號" in df.columns
        has_name = "品名"    in df.columns

        if has_erp:
            if has_name:
                # 組合顯示：「品號 - 品名」，空品名則只顯示品號
                def make_label(row):
                    pn   = str(row["ERP品號"]).strip() if pd.notna(row["ERP品號"]) else ""
                    name_ = str(row["品名"]).strip()   if pd.notna(row["品名"])    else ""
                    return f"{pn} - {name_}" if name_ else pn

                df = df.copy()
                df["_erp_label"] = df.apply(make_label, axis=1)
                erp_label_col = "_erp_label"
            else:
                erp_label_col = "ERP品號"
        else:
            erp_label_col = None

        # ── 篩選器（類型先選，ERP選項依類型聯動縮減）──────────────────────
        f1, f2 = st.columns(2)
        with f1:
            sel_brands = st.multiselect(
                "類型篩選", df[brand_col].dropna().unique(), key=f"{name}_brand"
            )

        # 先套用類型篩選，ERP 選項僅顯示該類型下的品號
        brand_filtered = df[df[brand_col].isin(sel_brands)] if sel_brands else df

        with f2:
            if erp_label_col:
                erp_options = brand_filtered[erp_label_col].dropna().unique().tolist()
                sel_erp_labels = st.multiselect(
                    "ERP品號篩選", erp_options, key=f"{name}_erp"
                )
            else:
                sel_erp_labels = []
                st.multiselect("ERP品號篩選", [], key=f"{name}_erp")

        filtered = brand_filtered.copy()
        if sel_erp_labels and erp_label_col:
            filtered = filtered[filtered[erp_label_col].isin(sel_erp_labels)]

        # ── KPI ─────────────────────────────────────────────────────────────
        st.markdown("#### 整體指標")
        kpi_row(filtered)
        st.markdown("---")

        # 自訂漸層配色
        bad_scale   = [[0, color_bad_low],   [1, color_bad_high]]
        reuse_scale = [[0, color_reuse_low],  [1, color_reuse_high]]

        # ── 圖表區 ──────────────────────────────────────────────────────────
        ch1, ch2 = st.columns(2)

        # 根據側邊欄模式決定數值標籤設定
        if label_mode.startswith("A"):
            txt_pos, txt_size, show_text = "outside",  12, True
        elif label_mode.startswith("B"):
            txt_pos, txt_size, show_text = "auto",     12, True
        elif label_mode.startswith("C"):
            txt_pos, txt_size, show_text = "outside",   9, True
        else:  # D
            txt_pos, txt_size, show_text = "outside",  12, False

        with ch1:
            fig = px.bar(
                filtered.sort_values("不良率(%)", ascending=False),
                x=brand_col, y="不良率(%)",
                color="不良率(%)",
                color_continuous_scale=bad_scale,
                title=f"{name}｜各類型不良率",
                text="不良率(%)" if show_text else None,
                labels={brand_col: "類型"},
            )
            if show_text:
                fig.update_traces(
                    texttemplate="%{text:.1f}%",
                    textposition=txt_pos,
                    textfont_size=txt_size,
                )
            fig.update_layout(
                coloraxis_showscale=False,
                height=420,
                yaxis_title="",
                annotations=[dict(
                    text="不良率 (%)",
                    x=-0.07, y=1.08,
                    xref="paper", yref="paper",
                    showarrow=False,
                    font=dict(size=13),
                )],
            )
            st.plotly_chart(fig, use_container_width=True)

        with ch2:
            avail_fault = [c for c in fault_cols if c in filtered.columns]
            fault_totals = filtered[avail_fault].sum()
            fault_totals = fault_totals[fault_totals > 0]
            if not fault_totals.empty:
                fig2 = px.pie(
                    values=fault_totals.values,
                    names=fault_totals.index,
                    title=f"{name}｜回廠原因分佈",
                )
                fig2.update_layout(height=420)
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("無回廠原因資料")

        ch3, ch4 = st.columns(2)

        with ch3:
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
                fig3.update_traces(
                    texttemplate="%{text:.1f}%",
                    textposition=txt_pos,
                    textfont_size=txt_size,
                )
            fig3.update_layout(
                coloraxis_showscale=False,
                height=420,
                yaxis_title="",
                annotations=[dict(
                    text="再使用率 (%)",
                    x=-0.07, y=1.08,
                    xref="paper", yref="paper",
                    showarrow=False,
                    font=dict(size=13),
                )],
            )
            st.plotly_chart(fig3, use_container_width=True)

        with ch4:
            disp_map = {
                "良品數":   "良品（再使用）",
                "不良品數": "不良品（維修/換貨）",
                "過保數":   "過保/報廢",
            }
            disp_vals = {v: int(filtered[k].sum()) for k, v in disp_map.items() if k in filtered.columns}
            disp_vals = {k: v for k, v in disp_vals.items() if v > 0}
            if disp_vals:
                fig4 = px.pie(
                    values=list(disp_vals.values()),
                    names=list(disp_vals.keys()),
                    hole=0.45,
                    title=f"{name}｜處置結果分佈",
                )
                fig4.update_layout(height=420)
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.info("無處置結果資料")

        # ── 上線量 vs 回廠量 ─────────────────────────────────────────────
        st.markdown("#### 上線量 vs 回廠量 對比")
        fig5 = go.Figure()
        fig5.add_trace(go.Bar(
            name="上線量", x=filtered[brand_col], y=filtered["上線量"],
            marker_color=color_online,
        ))
        fig5.add_trace(go.Bar(
            name="回廠量", x=filtered[brand_col], y=filtered["回廠量"],
            marker_color=color_return,
        ))
        fig5.update_layout(barmode="group", height=400, xaxis_title="類型", yaxis_title="數量")
        st.plotly_chart(fig5, use_container_width=True)

        st.markdown("---")

        # ── 詳細資料表（含廠牌小計） ─────────────────────────────────────
        st.markdown("#### 詳細資料")

        show_cols = [brand_col, "ERP品號", "品名", "上線量", "回廠量",
                     "良品數", "不良品數", "過保數",
                     "不良率(%)", "再使用率(%)", "過保率(%)"]
        show_cols = [c for c in show_cols if c in filtered.columns]
        num_cols  = ["上線量", "回廠量", "良品數", "不良品數", "過保數"]
        num_cols  = [c for c in num_cols if c in filtered.columns]

        # 每個廠牌加總 → 重新計算比率
        subtotals = (
            filtered[show_cols].groupby(brand_col, sort=False)[num_cols]
            .sum().reset_index()
        )
        subtotals["ERP品號"]     = "── 小計 ──"
        subtotals["品名"]        = ""
        subtotals["不良率(%)"]   = (subtotals["不良品數"] / subtotals["上線量"] * 100).fillna(0).round(1)
        subtotals["再使用率(%)"] = (subtotals["良品數"]   / subtotals["回廠量"] * 100).fillna(0).round(1)
        subtotals["過保率(%)"]   = (subtotals["過保數"]   / subtotals["上線量"] * 100).fillna(0).round(1)

        # 合併：每個品牌明細列 + 小計列，再加總計列
        frames = []
        for brand, grp in filtered[show_cols].groupby(brand_col, sort=False):
            frames.append(grp)
            frames.append(subtotals[subtotals[brand_col] == brand])

        # 全體總計
        total_row = {c: filtered[c].sum() for c in num_cols}
        total_row[brand_col] = "★ 總計"
        total_row["ERP品號"] = ""
        total_row["品名"]    = ""
        total_row["不良率(%)"]   = round(total_row["不良品數"] / total_row["上線量"] * 100, 1) if total_row.get("上線量") else 0
        total_row["再使用率(%)"] = round(total_row["良品數"]   / total_row["回廠量"] * 100, 1) if total_row.get("回廠量") else 0
        total_row["過保率(%)"]   = round(total_row["過保數"]   / total_row["上線量"] * 100, 1) if total_row.get("上線量") else 0
        frames.append(pd.DataFrame([total_row]))

        display_df = pd.concat(frames, ignore_index=True)[show_cols]

        # 標記小計/總計列（用於樣式）
        is_subtotal = display_df["ERP品號"] == "── 小計 ──"
        is_total    = display_df[brand_col]  == "★ 總計"

        # 將品牌欄位顯示名稱改為「類型」
        display_df = display_df.rename(columns={brand_col: "類型"})

        def highlight_rows(row):
            idx = row.name
            if is_total.iloc[idx]:
                return ["background-color: #1f4e79; color: white; font-weight: bold"] * len(row)
            if is_subtotal.iloc[idx]:
                return ["background-color: #d6e4f0; font-weight: bold"] * len(row)
            return [""] * len(row)

        styled = (
            display_df.style
            .apply(highlight_rows, axis=1)
            .format({
                "不良率(%)":   "{:.1f}",
                "再使用率(%)": "{:.1f}",
                "過保率(%)":   "{:.1f}",
            }, na_rep="")
        )

        st.dataframe(styled, use_container_width=True, height=500)

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
    st.info("請先上傳 Excel 檔案（需含「車機_彙整總覽」與「鏡頭_彙整總覽」工作表）")

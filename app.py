import streamlit as st
import pandas as pd
from io import BytesIO

# ===================== 页面设置 =====================
st.set_page_config(
    page_title="亚马逊数据工具箱",
    page_icon="📊",
    layout="wide"
)

# 自定义美观样式
st.markdown("""
<style>
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #2E8B57;
        text-align: center;
        margin-bottom: 30px;
    }
    .card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #2E8B57;
        color: white;
        font-size: 16px;
        padding: 10px 20px;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ===================== 标题 =====================
st.markdown('<p class="main-title">📦 亚马逊数据处理工具箱</p>', unsafe_allow_html=True)

# ===================== 选项卡 =====================
tab1, tab2 = st.tabs(["✅ 多文件合并", "✅ SIF关键词清洗"])

# ===================== 功能1：文件合并 =====================
with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📂 Excel / CSV 多文件合并")
    st.caption("支持批量上传文件，自动合并所有工作表，保留来源信息")

    uploaded_files = st.file_uploader(
        "选择要合并的文件",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="merge"
    )

    if uploaded_files and st.button("🚀 开始合并文件", type="primary"):
        all_data = []
        with st.spinner("正在处理..."):
            for file in uploaded_files:
                try:
                    if file.name.endswith((".xlsx", ".xls")):
                        excel = pd.ExcelFile(file)
                        for sheet in excel.sheet_names:
                            df = pd.read_excel(file, sheet_name=sheet)
                            df["来源文件"] = file.name
                            df["来源工作表"] = sheet
                            all_data.append(df)
                    else:
                        try:
                            df = pd.read_csv(file, encoding="utf-8")
                        except:
                            df = pd.read_csv(file, encoding="gbk")
                        df["来源文件"] = file.name
                        df["来源工作表"] = "CSV"
                        all_data.append(df)
                except:
                    st.warning(f"⚠️ {file.name} 读取失败")

        if all_data:
            res = pd.concat(all_data, ignore_index=True)
            st.success(f"✅ 合并完成！共 {len(res)} 行")
            st.dataframe(res.head(10), use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as w:
                res.to_excel(w, index=False)
            st.download_button("📥 下载合并结果", output.getvalue(), "合并结果.xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== 功能2：SIF 关键词清洗 =====================
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🎯 SIF 关键词数据清洗（亚马逊专用）")
    st.caption("自动：保留排名<30 + 去重最优排名 + 只保留需要的列")

    uploaded_file = st.file_uploader("上传SIF导出的表格", type=["xlsx", "csv"], key="sif")

    if uploaded_file and st.button("🧹 开始清洗数据", type="primary"):
        with st.spinner("正在清洗..."):
            # 读取（跳过第一行）
            try:
                if uploaded_file.name.endswith(".csv"):
                    df = pd.read_csv(uploaded_file, skiprows=1)
                else:
                    df = pd.read_excel(uploaded_file, skiprows=1)
            except:
                st.error("读取失败，请检查文件格式")
                st.stop()

            # 必须列
            must_cols = ["关键词", "自然排名"]
            for c in must_cols:
                if c not in df.columns:
                    st.error(f"缺少必要列：{c}")
                    st.stop()

            # 保留列
            KEEP = [
                "关键词", "翻译", "自然排名", "周搜索趋势",
                "关键词建议竞价（固定·精准)推荐", "关键词点击转化率"
            ]
            exist_keep = [c for c in KEEP if c in df.columns]
            df = df[exist_keep]

            # 排名 <30
            df["自然排名"] = pd.to_numeric(df["自然排名"], errors="coerce")
            df = df[df["自然排名"].notna() & (df["自然排名"] < 30)]

            # 清洗关键词
            df["关键词"] = df["关键词"].astype(str).str.strip().str.lower()
            df = df[df["关键词"] != ""]

            # 去重（保留最优排名）
            if len(df) > 0:
                best_idx = df.groupby("关键词")["自然排名"].idxmin()
                final = df.loc[best_idx].reset_index(drop=True)
            else:
                final = pd.DataFrame()

        if len(final) == 0:
            st.warning("⚠️ 没有符合条件的数据")
        else:
            st.success(f"✅ 清洗完成！原始：{len(df)} → 去重后：{len(final)}")
            st.dataframe(final, use_container_width=True)

            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                final.to_excel(w, index=False)
            st.download_button("📥 下载SIF清洗结果", out.getvalue(), "SIF清洗结果.xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== 底部 =====================
st.markdown("<br><div style='text-align:center; color:gray; font-size:14px'>亚马逊专用数据工具 | 安全·本地处理·不留档案</div>", unsafe_allow_html=True)

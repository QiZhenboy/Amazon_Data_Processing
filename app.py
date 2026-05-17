import streamlit as st
import pandas as pd
from io import BytesIO

# ===================== 页面设置 =====================
st.set_page_config(
    page_title="亚马逊数据处理工具箱",
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

# ===================== 功能2：SIF 关键词清洗（升级版） =====================
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🎯 SIF 关键词数据清洗（亚马逊专用）")
    st.caption("自定义：排名阈值 + 保留列；自动去重并保留最优排名")

    # 1. 上传文件
    uploaded_file = st.file_uploader("上传SIF导出的表格", type=["xlsx", "csv"], key="sif")

    if uploaded_file:
        # 2. 先读取文件，获取所有列名，给用户选择
        with st.spinner("正在读取文件，获取列名..."):
            try:
                if uploaded_file.name.endswith(".csv"):
                    temp_df = pd.read_csv(uploaded_file, skiprows=1)
                else:
                    temp_df = pd.read_excel(uploaded_file, skiprows=1)
            except Exception as e:
                st.error(f"读取文件失败：{str(e)}")
                st.stop()

        # 3. 用户配置区域
        st.subheader("⚙️ 清洗设置")
        # 排名阈值
        rank_threshold = st.number_input(
            "保留自然排名小于多少？",
            min_value=1,
            max_value=1000,
            value=30,
            help="例如填30，会只保留自然排名1-29的关键词"
        )

        # 自定义保留列（默认选中常用列）
        default_cols = [
            "关键词", "翻译", "自然排名", "周搜索趋势",
            "关键词建议竞价（固定·精准)推荐", "关键词点击转化率"
        ]
        # 只保留文件里实际存在的列作为默认选中项
        valid_default_cols = [col for col in default_cols if col in temp_df.columns]
        keep_cols = st.multiselect(
            "选择要保留的列",
            options=temp_df.columns.tolist(),
            default=valid_default_cols,
            help="你可以根据需要勾选要保留的列，不勾选的列会被删除"
        )

        # 必须列检查（关键词、自然排名必须保留）
        if "关键词" not in keep_cols or "自然排名" not in keep_cols:
            st.warning("⚠️ 请务必勾选「关键词」和「自然排名」，否则无法去重和筛选！")
            st.stop()

        # 4. 开始清洗按钮
        if st.button("🧹 开始清洗数据", type="primary"):
            with st.spinner("正在清洗数据..."):
                df = temp_df.copy()

                # 只保留用户选择的列
                df = df[keep_cols]

                # 筛选排名 < 用户设置的阈值
                df["自然排名"] = pd.to_numeric(df["自然排名"], errors="coerce")
                df = df[df["自然排名"].notna() & (df["自然排名"] < rank_threshold)]

                # 清洗关键词
                df["关键词"] = df["关键词"].astype(str).str.strip().str.lower()
                df = df[df["关键词"] != ""]

                # 去重：保留排名最好（数字最小）的一条
                if len(df) > 0:
                    best_idx = df.groupby("关键词")["自然排名"].idxmin()
                    final_df = df.loc[best_idx].reset_index(drop=True)
                else:
                    final_df = pd.DataFrame()

            # 5. 显示结果
            if len(final_df) == 0:
                st.warning("⚠️ 没有符合条件的数据，请检查排名阈值或文件内容")
            else:
                st.success(f"✅ 清洗完成！原始有效数据：{len(df)} → 去重后：{len(final_df)}")
                st.dataframe(final_df, use_container_width=True)

                # 6. 下载结果
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    final_df.to_excel(writer, index=False)
                st.download_button(
                    label="📥 下载SIF清洗结果",
                    data=output.getvalue(),
                    file_name="SIF清洗结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.markdown('</div>', unsafe_allow_html=True)

# ===================== 底部 =====================
st.markdown(
    "<br><div style='text-align:center; color:gray; font-size:14px'>亚马逊专用数据工具 | 安全·本地处理·不留档案</div>",
    unsafe_allow_html=True
)

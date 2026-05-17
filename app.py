import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 页面设置
st.set_page_config(
    page_title="万能表格合并工具",
    page_icon="📊",
    layout="wide"
)

# 标题
st.title("📊 万能表格合并工具")
st.markdown("支持 **Excel(.xlsx/.xls)** + **CSV** 多文件批量合并，自动保留来源信息")

# 上传文件
uploaded_files = st.file_uploader(
    "⬇️ 拖放或选择要合并的文件",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

# 开始合并按钮
if uploaded_files and st.button("🚀 开始合并"):
    all_data = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    # 遍历所有上传的文件
    for i, file in enumerate(uploaded_files):
        status_text.text(f"正在处理：{file.name}")
        progress = (i + 1) / len(uploaded_files)
        progress_bar.progress(progress)

        try:
            # 处理 Excel 文件
            if file.name.endswith((".xlsx", ".xls")):
                excel_file = pd.ExcelFile(file)
                for sheet in excel_file.sheet_names:
                    df = pd.read_excel(file, sheet_name=sheet)
                    df["来源文件"] = file.name
                    df["来源工作表"] = sheet
                    all_data.append(df)

            # 处理 CSV 文件
            elif file.name.endswith(".csv"):
                try:
                    df = pd.read_csv(file, encoding="utf-8")
                except:
                    df = pd.read_csv(file, encoding="gbk")
                df["来源文件"] = file.name
                df["来源工作表"] = "CSV文件"
                all_data.append(df)

        except Exception as e:
            st.warning(f"⚠️ {file.name} 处理失败：{str(e)}")

    # 合并完成
    if all_data:
        merged = pd.concat(all_data, ignore_index=True)
        status_text.text("✅ 合并完成！")

        # 展示预览
        st.subheader("📋 合并结果预览")
        st.dataframe(merged.head(10), use_container_width=True)
        st.success(f"合并成功！共 {len(merged)} 行数据")

        # 生成下载文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False)
        output.seek(0)

        # 下载按钮
        st.download_button(
            label="📥 下载合并结果.xlsx",
            data=output,
            file_name="合并结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("❌ 没有成功读取到数据")

# 使用说明
with st.expander("📖 使用说明"):
    st.markdown("""
    1. 支持格式：.xlsx / .xls / .csv
    2. 可同时上传**多个文件**
    3. Excel 会自动合并**所有工作表**
    4. 最后一列会标记**来源文件+工作表**
    5. 手机、电脑、平板均可使用
    """)
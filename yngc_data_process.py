import base64
import pandas as pd
import streamlit as st
from io import BytesIO


# 设置页面配置
st.set_page_config(
    page_title="YNGC数据处理工具",  # 设置应用名称
    page_icon=":material/home:",   # 设置应用图标
)

# 将 Pandas DataFrame 对象转换为 Excel 文件格式的字节流
def to_excel(df):
    output = BytesIO()
    # 使用 openpyxl 作为引擎
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# 创建自定义下载按钮的函数
def create_download_button(df, label):
    excel_file = to_excel(df)
    b64 = base64.b64encode(excel_file).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{label}.xlsx">下载 {label} 数据文件</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.header('YNGC数据处理工具', divider='rainbow')
st.subheader('1.批次节点数据填写',divider='grey')
cx_node = pd.DataFrame(
    [
        {"产线": "1", "批次节点": None },
        {"产线": "2", "批次节点": None },
        {"产线": "3", "批次节点": None },
        {"产线": "4", "批次节点": None },
        {"产线": "5", "批次节点": None },
        {"产线": "6", "批次节点": None },
        {"产线": "7", "批次节点": None },
        {"产线": "8", "批次节点": None },
        {"产线": "9", "批次节点": None },
        {"产线": "10", "批次节点": None }
    ]
)
edited_df = st.data_editor(cx_node, hide_index=True, width=600)

st.subheader('2.模版数据文件上传',divider='grey')
# 文件上传
uploaded_file = st.file_uploader("请选择一个Excel文件(.xlsx格式)上传", type=["xlsx"])

if uploaded_file is not None:
    # 读取文件
    upload_df = pd.read_excel(uploaded_file,header=1)
    # 处理文件
    df_raw = upload_df.loc[1:,:]
    cols = ['序号','样品批号', '下样时间', '外观\n','乙烯基链节摩尔分数\n%', '挥发分（150℃，3h）\n%wt', '相对分子量x1e4\n']
    df_raw = df_raw.drop([df_raw.index[0], df_raw.index[-1]])
    tdf = df_raw[cols].copy()  # 创建副本以避免警告
    # 数据类型转换
    tdf['样品批号'] = tdf['样品批号'].astype(int).astype(str)
    tdf['下样时间'] = pd.to_datetime(tdf['下样时间'])
    tdf['下样时间'] = tdf['下样时间'].dt.date
    tdf['下样时间'] = tdf['下样时间'].astype(str)
    
    # 提取关键信息
    # 提取批次节点
    tdf['批次号'] = tdf['样品批号'].str[-4:].astype(int)
    # 根据字符串长度提取中间的一位或两位
    tdf['产线'] = tdf['样品批号'].apply(lambda x: x[6] if len(x) == 11 else x[6:8])

    # 记录本次处理后的批次节点数据
    pc_max = tdf.groupby('产线')['批次号'].max().reset_index()
    pc_max['产线'] = pc_max['产线'].astype(int)
    pc_max = pc_max.sort_values(by=['产线'])

    cdf = edited_df
    # 筛选出批次节点不为None的行
    cdf_filtered = cdf[cdf['批次节点'].notna()].copy()
    cdf_filtered['批次节点'] = cdf_filtered['批次节点'].astype(int)
    # 初始化结果DataFrame
    result_df = pd.DataFrame(columns=tdf.columns)

    # 遍历cdf_filtered中的每一行
    for index, row in cdf_filtered.iterrows():
        # 筛选出tdf中产线与cdf产线对应且批次号大于批次节点的数据
        filtered_rows = tdf[(tdf['产线'] == row['产线']) & (tdf['批次号'] > row['批次节点'])]
        # 将结果添加到result_df中
        result_df = pd.concat([result_df, filtered_rows])
    
    # 重置索引
    result_df.reset_index(drop=True, inplace=True)
    result_df = result_df[cols]
    
    df_red = result_df[result_df['乙烯基链节摩尔分数\n%'] == '0.04']
    df_green = result_df[result_df['乙烯基链节摩尔分数\n%'] == '0.08']
    df_yellow = result_df[result_df['乙烯基链节摩尔分数\n%'] == '0.16']
    df_blue = result_df[result_df['乙烯基链节摩尔分数\n%'] =='0.23' ]
    
    st.subheader("3.Excel文件下载",divider='grey')

    # 创建一个 2x2 的布局
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("MVQ110-0_红色", divider='red')
        create_download_button(df_red, "MVQ110-0_红色")

        st.subheader("MVQ110-1_绿色", divider='green')
        create_download_button(df_green, "MVQ110-1_绿色")

    with col2:
        st.subheader("MVQ110-2_黄色", divider='orange')
        create_download_button(df_yellow, "MVQ110-2_黄色")

        st.subheader("MVQ110-3_蓝色", divider='blue')
        create_download_button(df_blue, "MVQ110-3_蓝色")
    st.subheader("各产线批次节点数据", divider='grey')
    create_download_button(pc_max, "当前各产线批次节点")
import streamlit as st
import pandas as pd
import os

# 页面基本配置
st.set_page_config(page_title="数据看板", page_icon="📈", layout="wide")

# 自定义CSS
st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        padding-top: 10px;
        padding-bottom: 10px;
        font-size: 16px;
        font-weight: 600;
    }
    /* 针对汇总 HTML 表格的美化及居中 */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-family: sans-serif;
        font-size: 14px;
        margin-bottom: 2rem;
    }
    .custom-table th, .custom-table td {
        border: 1px solid #e0e0e0;
        padding: 10px 12px;
        text-align: center !important;
        vertical-align: middle !important;
    }
    .custom-table thead th {
        background-color: #f7f7f9;
        font-weight: 600;
        color: #31333F;
    }
    /* 隐藏HTML表格右上角的索引名称栏 */
    .custom-table thead tr:last-child th {
        border-top: none;
    }
    </style>
""", unsafe_allow_html=True)

# 金山文档 (WPS) 在线分享 ID
WPS_FILE_ID = "cavvuRYktATy"

# 数据源候选路径（本地兜底）
def get_data_file_path():
    candidate_paths = [
        r"C:\Users\18501\Desktop\SA旗舰店及双高体验馆&生活馆.xlsx",
        os.path.expanduser(r"~\Desktop\SA旗舰店及双高体验馆&生活馆.xlsx"),
        os.path.join(os.path.dirname(__file__), "SA旗舰店及双高体验馆&生活馆.xlsx")
    ]
    for p in candidate_paths:
        if os.path.exists(p):
            return p
    return candidate_paths[0]

def parse_excel_stream(excel_bytes):
    import openpyxl
    import io
    
    wb = openpyxl.load_workbook(excel_bytes, data_only=True)
    ws_summary = wb['汇总']
    ws_store = wb['门店分析']
    hidden_summary_rows = [r for r, dim in ws_summary.row_dimensions.items() if dim.hidden]
    hidden_store_rows = [r for r, dim in ws_store.row_dimensions.items() if dim.hidden]
    wb.close()

    excel_bytes.seek(0)
    df_summary = pd.read_excel(excel_bytes, sheet_name='汇总', header=1)
    excel_bytes.seek(0)
    df_store = pd.read_excel(excel_bytes, sheet_name='门店分析', header=1)
    
    # 提取用户在截图中所呈现的 21 列 (索引 0~18, 24, 68)
    keep_indices = list(range(0, 19)) + [24, 68]
    valid_indices = [i for i in keep_indices if i < len(df_store.columns)]
    df_store = df_store.iloc[:, valid_indices].copy()
    
    # Excel行号是 1-based，且 header=1 意味着第1行是标题，第2行是表头，第3行是数据的第一行(对应Pandas的index 0)
    idx_drop_summary = [r - 3 for r in hidden_summary_rows if (r - 3) in df_summary.index]
    df_summary = df_summary.drop(index=idx_drop_summary)
    
    idx_drop_store = [r - 3 for r in hidden_store_rows if (r - 3) in df_store.index]
    df_store = df_store.drop(index=idx_drop_store)
    
    # 清理Unnamed列
    df_summary = df_summary.loc[:, ~df_summary.columns.str.contains('^Unnamed')]
    
    # 门店分析层面：只保留所属运中为黑吉的门店
    if '门店所属运中' in df_store.columns:
        df_store = df_store[df_store['门店所属运中'].astype(str).str.strip() == '黑吉'].copy()

    # 规范化“是否活跃”列，将“/”或空值统一表示为“不活跃”
    if '是否活跃' in df_store.columns:
        df_store['是否活跃'] = df_store['是否活跃'].replace({'/': '不活跃', '-': '不活跃'}).fillna('不活跃')

    # 百分比格式化函数，保留整数
    def format_pct(x):
        if pd.isna(x) or str(x).strip() in ["", "-", "nan", "none", "None", "NULL", "null"]:
            return ""
        try:
            return f"{float(x) * 100:.0f}%"
        except (ValueError, TypeError):
            return x

    # 整数格式化函数（确保门店数量和活跃数量不保留一位小数）
    def format_int(x):
        if pd.isna(x) or str(x).strip() in ["", "-", "nan", "none", "None", "NULL", "null"]:
            return ""
        try:
            return str(int(round(float(x))))
        except (ValueError, TypeError):
            return x

    # 向下填充汇总表的 地区 和 业务 字段，以便生成 MultiIndex 从而实现行合并居中
    if '地区' in df_summary.columns and '业务' in df_summary.columns:
        df_summary['地区'] = df_summary['地区'].ffill()
        summary_mask = df_summary['地区'].astype(str).str.contains('计', na=False)
        df_summary['业务'] = df_summary['业务'].ffill().where(~summary_mask, "")
        if '所属督导' in df_summary.columns:
            df_summary['所属督导'] = df_summary['所属督导'].fillna("")

    # 汇总层面：格式化占比与数量（数量显示为纯整数）
    for col in df_summary.columns:
        if '占比' in col or '同比' in col:
            df_summary[col] = df_summary[col].apply(format_pct)
        elif any(kw in col for kw in ['数量', '店数', '户数']):
            df_summary[col] = df_summary[col].apply(format_int)

    return df_summary, df_store

@st.cache_data(ttl=60)
def load_data():
    import requests
    import io
    
    # 优先尝试从金山文档 (WPS) 在线实时拉取
    try:
        api_url = f"https://www.kdocs.cn/api/office/file/{WPS_FILE_ID}/download"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        resp = requests.get(api_url, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            download_url = data.get("download_url")
            if download_url:
                excel_resp = requests.get(download_url, headers=headers, timeout=30)
                if excel_resp.status_code == 200:
                    excel_bytes = io.BytesIO(excel_resp.content)
                    return parse_excel_stream(excel_bytes), "在线金山文档"
    except Exception:
        pass

    # 若无法连接网络则读取本地文件兜底
    local_path = get_data_file_path()
    if os.path.exists(local_path):
        import io
        with open(local_path, "rb") as f:
            return parse_excel_stream(io.BytesIO(f.read())), "本地备份文件"

    raise RuntimeError("无法从金山文档或本地找到数据源文件。")

def main():
    col_t1, col_t2 = st.columns([4, 1])
    with col_t1:
        st.title("📈 黑吉SAB旗舰店（含双高）Q3商机管理客资活跃情况")
    with col_t2:
        st.write("")
        if st.button("🔄 同步金山文档最新数据", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    st.markdown("---")

    try:
        (df_summary, df_store), source_label = load_data()
    except Exception as e:
        st.error(f"读取数据失败: {e}")
        return

    # 选项卡和表头名称完全与原始表格的 Sheet 名称保持一致
    tab1, tab2 = st.tabs(["📊 汇总", "🏪 门店分析"])

    with tab1:
        st.subheader("汇总")
        
        # 将前三列提取为 index，利用 pandas to_html 的 sparsify=True 来合并单元格
        idx_cols = [c for c in ['地区', '业务', '所属督导'] if c in df_summary.columns]
        if idx_cols:
            df_summary_disp = df_summary.set_index(idx_cols)
        else:
            df_summary_disp = df_summary
            
        html_table = df_summary_disp.to_html(classes="custom-table", escape=False, sparsify=True)
        
        # 修复 pandas to_html 默认生成的多行交错表头（这会导致页面显示错位/乱码）
        import re
        thead_html = "<thead><tr>"
        colgroup_html = "<colgroup>"
        for i, col in enumerate(df_summary.columns):
            # 将第一列（地区）的宽度调小固定
            if i == 0:
                colgroup_html += "<col style='width: 80px;'>"
            else:
                colgroup_html += "<col>"
                
            # 仅在显示阶段去除 .1, .2 等后缀，不影响底层数据结构
            clean_col = re.sub(r'\.\d+$', '', str(col))
            thead_html += f"<th>{clean_col}</th>"
        thead_html += "</tr></thead>"
        colgroup_html += "</colgroup>"
        
        # 将构造好的 colgroup 和 规整表头 替换掉原来的 <thead>
        html_table = re.sub(r'<thead>.*?</thead>', colgroup_html + thead_html, html_table, flags=re.DOTALL)
        
        # 合并“总计”或“合计”所在行的空白 <th> 单元格
        def merge_th(match):
            full_str = match.group(0)
            col_span = full_str.count('<th')
            inner_text = re.search(r'<th>(.*?)</th>', match.group(1)).group(1)
            return f'<th colspan="{col_span}">{inner_text}</th>'
            
        html_table = re.sub(r'(<th>[^<]*(?:总计|合计)[^<]*</th>)(?:\s*<th></th>)+', merge_th, html_table)
        
        # 将“备注”所在的最后一行完全合并为一个 <td> 单元格
        total_cols = len(df_summary.columns)
        def merge_note(match):
            inner_text = match.group(1)
            # 将字面量的 \n 或换行符替换为 HTML 的 <br>，字号调整得更小更精致，文字靠左对齐
            formatted_text = inner_text.replace('\\n', '<br>').replace('\n', '<br>')
            return f'<tr><td colspan="{total_cols}" style="text-align: left !important; font-size: 11px; color: #666; padding: 8px 12px; line-height: 1.6;">{formatted_text}</td></tr>'
            
        html_table = re.sub(r'<tr>\s*<th>(备注[^<]*)</th>.*?</tr>', merge_note, html_table, flags=re.DOTALL)
        
        st.markdown(html_table, unsafe_allow_html=True)

    with tab2:
        st.subheader("门店分析 (条件筛选)")
        
        filter_cols = ['所属区域', '门店最新等级', '是否需要跟进', '是否活跃']
        available_filters = [c for c in filter_cols if c in df_store.columns]
        
        cols = st.columns(len(available_filters) if available_filters else 1)
        
        selected_filters = {}
        for i, col_name in enumerate(available_filters):
            # 取唯一值，并剔除无效选项（如 0, -, 空白等）
            raw_vals = df_store[col_name].dropna().unique()
            unique_vals = []
            for v in raw_vals:
                s_val = str(v).strip()
                if s_val not in ["", "0", "-", "nan", "None", "NULL", "null"]:
                    if s_val not in unique_vals:
                        unique_vals.append(s_val)
            
            # 对各选项进行人性化排序
            if col_name == '是否活跃':
                unique_vals = [x for x in ['活跃', '不活跃'] if x in unique_vals] or unique_vals
            elif col_name == '是否需要跟进':
                unique_vals = [x for x in ['需要', '不需要'] if x in unique_vals] or unique_vals
            elif col_name == '门店最新等级':
                order = ['S', 'A', 'B', 'C', 'D']
                unique_vals = sorted(unique_vals, key=lambda x: order.index(x) if x in order else 99)
            
            with cols[i]:
                selected = st.selectbox(f"筛选: {col_name}", options=["全部"] + unique_vals)
                if selected != "全部":
                    selected_filters[col_name] = selected
                    
        # 应用筛选
        filtered_df = df_store.copy()
        for col, selected_val in selected_filters.items():
            filtered_df = filtered_df[filtered_df[col].astype(str) == selected_val]
            
        st.markdown(f"**当前筛选结果:** 共查询到 `{len(filtered_df)}` 家门店")
        
        # 使用 column_config 在界面显示时去除表头后缀 (如 .1 等)
        import re
        col_cfg = {c: st.column_config.Column(label=re.sub(r'\.\d+$', '', str(c))) for c in filtered_df.columns}
        st.dataframe(filtered_df, use_container_width=True, hide_index=True, column_config=col_cfg)

if __name__ == "__main__":
    main()


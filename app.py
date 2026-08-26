import streamlit as st
import pandas as pd
import os
import re
import io
import requests

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

    /* 督导分析专用样式 */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #eef0f4;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    .metric-title {
        font-size: 13px;
        color: #666;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 700;
        color: #1f1f1f;
    }

    .sup-table {
        width: 100%;
        border-collapse: collapse;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        font-size: 13px;
        margin-top: 12px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        border-radius: 8px;
        overflow: hidden;
    }
    .sup-table th {
        background-color: #1e3d59;
        color: #ffffff;
        font-weight: 600;
        text-align: center;
        padding: 11px 14px;
        border: 1px solid #173046;
    }
    .sup-table td {
        padding: 10px 12px;
        text-align: center;
        border: 1px solid #e8e8e8;
        color: #333333;
        vertical-align: middle;
    }
    .sup-table tr:hover {
        background-color: #f5f9ff;
    }
    /* 后10名标红预警行样式 */
    .sup-warning-row {
        background-color: #fff1f0 !important;
    }
    .sup-warning-row td {
        border-color: #ffccc7 !important;
    }
    .sup-warning-row:hover {
        background-color: #ffe6e6 !important;
    }
    .badge-warning {
        background-color: #ff4d4f;
        color: #ffffff;
        font-size: 12px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 12px;
        display: inline-block;
        box-shadow: 0 1px 2px rgba(255, 77, 79, 0.2);
    }
    .badge-normal {
        background-color: #e6f7ff;
        color: #1890ff;
        font-size: 12px;
        font-weight: 600;
        padding: 3px 10px;
        border-radius: 12px;
        display: inline-block;
    }
    .badge-rank1 {
        background-color: #fffbe6;
        color: #d48806;
        font-size: 12px;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 12px;
        border: 1px solid #ffe58f;
        display: inline-block;
    }
    .gap-neg {
        color: #cf1322;
        font-weight: 700;
    }
    .gap-pos {
        color: #389e0d;
        font-weight: 700;
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

def extract_hidden_rows_fast(excel_bytes):
    import zipfile
    import xml.etree.ElementTree as ET
    
    hidden_summary = []
    hidden_store = []
    try:
        with zipfile.ZipFile(excel_bytes, 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            root_wb = ET.fromstring(wb_xml)
            namespaces = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            rels_xml = z.read('xl/_rels/workbook.xml.rels')
            root_rels = ET.fromstring(rels_xml)
            rel_ns = {'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            rel_map = {r.attrib['Id']: r.attrib['Target'] for r in root_rels.findall('pkg:Relationship', rel_ns)}
            
            sheet_files = {}
            for sheet in root_wb.findall('.//main:sheet', namespaces):
                name = sheet.attrib.get('name')
                rId = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = rel_map.get(rId)
                if name and target:
                    sheet_files[name] = 'xl/' + target if not target.startswith('xl/') else target

            if '汇总' in sheet_files:
                root_sheet = ET.fromstring(z.read(sheet_files['汇总']))
                for row in root_sheet.findall('.//main:row', namespaces):
                    if row.attrib.get('hidden') == '1':
                        hidden_summary.append(int(row.attrib['r']))
                        
            if '门店分析' in sheet_files:
                root_sheet = ET.fromstring(z.read(sheet_files['门店分析']))
                for row in root_sheet.findall('.//main:row', namespaces):
                    if row.attrib.get('hidden') == '1':
                        hidden_store.append(int(row.attrib['r']))
    except Exception:
        pass
    return hidden_summary, hidden_store

def parse_excel_stream(excel_bytes):
    excel_bytes.seek(0)
    with pd.ExcelFile(excel_bytes, engine='openpyxl') as ef:
        df_summary = pd.read_excel(ef, sheet_name='汇总', header=1)
        df_store = pd.read_excel(ef, sheet_name='门店分析', header=1)
    
    # 提取用户在截图中所呈现的 21 列 (索引 0~18, 24, 68)
    keep_indices = list(range(0, 19)) + [24, 68]
    valid_indices = [i for i in keep_indices if i < len(df_store.columns)]
    df_store = df_store.iloc[:, valid_indices].copy()
    
    # 清理Unnamed列与空行
    df_summary = df_summary.loc[:, ~df_summary.columns.str.contains('^Unnamed')].dropna(how='all')
    df_store = df_store.loc[:, ~df_store.columns.str.contains('^Unnamed')].dropna(subset=['门店编码'])
    
    # 门店分析层面：只保留所属运中为黑吉，且门店类型非专卖店（旗舰店、星级生活馆、智感体验馆共146家，含S/A/B/C/D全部等级）
    if '门店所属运中' in df_store.columns:
        df_store = df_store[df_store['门店所属运中'].astype(str).str.strip() == '黑吉'].copy()
        
    if '门店类型' in df_store.columns:
        df_store = df_store[df_store['门店类型'].astype(str).str.strip() != '专卖店'].copy()

    # 规范化“是否活跃”列，将“/”或空值统一表示为“不活跃”
    if '是否活跃' in df_store.columns:
        df_store['是否活跃'] = df_store['是否活跃'].replace({'/': '不活跃', '-': '不活跃'}).fillna('不活跃')

    # 向下填充汇总表的 地区 和 业务 字段
    if '地区' in df_summary.columns and '业务' in df_summary.columns:
        df_summary['地区'] = df_summary['地区'].ffill()
        summary_mask = df_summary['地区'].astype(str).str.contains('计', na=False)
        df_summary['业务'] = df_summary['业务'].ffill().where(~summary_mask, "")
        if '所属督导' in df_summary.columns:
            df_summary['所属督导'] = df_summary['所属督导'].fillna("")

    # --- 构造督导分析专属数据表 (在汇总表进行文本格式化之前提取纯数值) ---
    df_sup_clean = df_summary[~df_summary['所属督导'].isna() & (df_summary['所属督导'].astype(str).str.strip() != '')].copy()
    df_sup_clean = df_sup_clean[~df_sup_clean['所属督导'].astype(str).str.contains('计|总|合')].copy()

    # 计算各督导管辖门店的客资缺口总额
    df_store_for_gap = df_store.copy()
    if '缺口' in df_store_for_gap.columns and '所属督导' in df_store_for_gap.columns:
        df_store_for_gap['缺口_数值'] = pd.to_numeric(df_store_for_gap['缺口'], errors='coerce').fillna(0)
        gap_dict = df_store_for_gap.groupby('所属督导')['缺口_数值'].sum().to_dict()
    else:
        gap_dict = {}

    df_supervisor = pd.DataFrame()
    df_supervisor['所属督导'] = df_sup_clean['所属督导']
    df_supervisor['地区'] = df_sup_clean['地区']
    df_supervisor['业务'] = df_sup_clean['业务']
    df_supervisor['门店数量'] = pd.to_numeric(df_sup_clean['门店数量（SAB）'], errors='coerce').fillna(0).astype(int)
    df_supervisor['活跃数量'] = pd.to_numeric(df_sup_clean['活跃数量'], errors='coerce').fillna(0).astype(int)
    df_supervisor['活跃率_数值'] = pd.to_numeric(df_sup_clean['活跃门店占比'], errors='coerce').fillna(0)
    df_supervisor['客资缺口总额'] = df_supervisor['所属督导'].map(gap_dict).fillna(0).astype(int)

    # 排序：活跃率降序，门店数量降序
    df_supervisor = df_supervisor.sort_values(by=['活跃率_数值', '门店数量'], ascending=[False, False]).reset_index(drop=True)
    df_supervisor['排名'] = range(1, len(df_supervisor) + 1)
    
    # 标记后10名预警 (倒数10名)
    total_sups = len(df_supervisor)
    bottom_10_start = max(1, total_sups - 9)
    df_supervisor['是否后10名预警'] = df_supervisor['排名'] >= bottom_10_start
    df_supervisor['活跃率'] = df_supervisor['活跃率_数值'].apply(lambda x: f"{x*100:.0f}%")

    # --- 汇总表文本格式化 (供Tab 1渲染) ---
    def format_pct(x):
        if pd.isna(x) or str(x).strip() in ["", "-", "nan", "none", "None", "NULL", "null"]:
            return ""
        try:
            return f"{float(x) * 100:.0f}%"
        except (ValueError, TypeError):
            return x

    def format_int(x):
        if pd.isna(x) or str(x).strip() in ["", "-", "nan", "none", "None", "NULL", "null"]:
            return ""
        try:
            return str(int(round(float(x))))
        except (ValueError, TypeError):
            return x

    for col in df_summary.columns:
        if '占比' in col or '同比' in col:
            df_summary[col] = df_summary[col].apply(format_pct)
        elif any(kw in col for kw in ['数量', '店数', '户数']):
            df_summary[col] = df_summary[col].apply(format_int)

    return df_summary, df_store, df_supervisor

@st.cache_data(ttl=180)
def load_data():
    import requests
    import io
    
    # 优先尝试从金山文档 (WPS) 在线实时拉取
    try:
        api_url = f"https://www.kdocs.cn/api/office/file/{WPS_FILE_ID}/download"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        with requests.Session() as s:
            s.trust_env = False
            resp = s.get(api_url, headers=headers, timeout=8)
            if resp.status_code == 200:
                data = resp.json()
                download_url = data.get("download_url")
                if download_url:
                    excel_resp = s.get(download_url, headers=headers, timeout=20)
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
        (df_summary, df_store, df_supervisor), source_label = load_data()
    except Exception as e:
        st.error(f"读取数据失败: {e}")
        return

    # 3个 Tab 页面
    tab1, tab2, tab3 = st.tabs(["📊 汇总", "🏪 门店分析", "👨‍💼 督导分析"])

    with tab1:
        st.subheader("汇总")
        
        # 将前三列提取为 index，利用 pandas to_html 的 sparsify=True 来合并单元格
        idx_cols = [c for c in ['地区', '业务', '所属督导'] if c in df_summary.columns]
        if idx_cols:
            df_summary_disp = df_summary.set_index(idx_cols)
        else:
            df_summary_disp = df_summary
            
        html_table = df_summary_disp.to_html(classes="custom-table", escape=False, sparsify=True)
        
        # 修复 pandas to_html 默认生成的多行交错表头
        thead_html = "<thead><tr>"
        colgroup_html = "<colgroup>"
        for i, col in enumerate(df_summary.columns):
            if i == 0:
                colgroup_html += "<col style='width: 80px;'>"
            else:
                colgroup_html += "<col>"
                
            clean_col = re.sub(r'\.\d+$', '', str(col))
            thead_html += f"<th>{clean_col}</th>"
        thead_html += "</tr></thead>"
        colgroup_html += "</colgroup>"
        
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
            formatted_text = inner_text.replace('\\n', '<br>').replace('\n', '<br>')
            return f'<tr><td colspan="{total_cols}" style="text-align: left !important; font-size: 11px; color: #666; padding: 8px 12px; line-height: 1.6;">{formatted_text}</td></tr>'
            
        html_table = re.sub(r'<tr>\s*<th>(备注[^<]*)</th>.*?</tr>', merge_note, html_table, flags=re.DOTALL)
        
        st.markdown(html_table, unsafe_allow_html=True)

    with tab2:
        st.subheader("门店分析 (条件筛选)")
        
        filter_cols = ['所属区域', '最新冻结等级', '门店最新等级', '是否需要跟进', '是否活跃']
        available_filters = [c for c in filter_cols if c in df_store.columns]
        
        cols = st.columns(len(available_filters) if available_filters else 1)
        
        selected_filters = {}
        for i, col_name in enumerate(available_filters):
            # 取唯一值，并剔除无效选项
            raw_vals = df_store[col_name].dropna().unique()
            unique_vals = []
            for v in raw_vals:
                s_val = str(v).strip()
                if s_val not in ["", "0", "-", "nan", "None", "NULL", "null"]:
                    if s_val not in unique_vals:
                        unique_vals.append(s_val)
            
            # 对各选项进行人性化排序并确保固定档位完整
            if col_name == '是否活跃':
                unique_vals = ['活跃', '不活跃']
            elif col_name == '是否需要跟进':
                unique_vals = ['需要', '不需要']
            elif col_name in ['门店最新等级', '最新冻结等级']:
                order = ['S', 'A', 'B', 'C', 'D']
                unique_vals = order + [x for x in unique_vals if x not in order]
            
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
        col_cfg = {c: st.column_config.Column(label=re.sub(r'\.\d+$', '', str(c))) for c in filtered_df.columns}
        st.dataframe(filtered_df, use_container_width=True, hide_index=True, column_config=col_cfg)

    with tab3:
        st.subheader("督导分析 (活跃率排名 & 后10名标红预警)")

        # 统计卡片指标
        total_sup_count = len(df_supervisor)
        warning_sup_count = df_supervisor['是否后10名预警'].sum()
        avg_active_rate = df_supervisor['活跃率_数值'].mean()
        warn_df = df_supervisor[df_supervisor['是否后10名预警']]
        warn_total_gap = warn_df['客资缺口总额'].sum()

        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">督导总人数</div>
                    <div class="metric-value">{total_sup_count} 人</div>
                </div>
            """, unsafe_allow_html=True)
        with m_col2:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">⚠️ 后10名预警人数</div>
                    <div class="metric-value" style="color: #cf1322;">{warning_sup_count} 人</div>
                </div>
            """, unsafe_allow_html=True)
        with m_col3:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">督导平均活跃率</div>
                    <div class="metric-value">{avg_active_rate*100:.1f}%</div>
                </div>
            """, unsafe_allow_html=True)
        with m_col4:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">预警督导客资缺口总额</div>
                    <div class="metric-value" style="color: #cf1322;">{warn_total_gap:,}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # 筛选模式
        filter_mode = st.radio(
            "查看范围:",
            options=["全部督导排名 (23人)", "⚠️ 仅看后10名预警督导 (10人)"],
            horizontal=True
        )

        disp_sup = df_supervisor.copy()
        if "仅看后10名" in filter_mode:
            disp_sup = disp_sup[disp_sup['是否后10名预警']].copy()

        # 生成督导分析的 HTML 响应式表格
        sup_html = """
        <table class="sup-table">
            <thead>
                <tr>
                    <th style="width: 80px;">排名</th>
                    <th style="width: 100px;">所属督导</th>
                    <th style="width: 110px;">地区</th>
                    <th style="width: 100px;">业务</th>
                    <th style="width: 90px;">门店数量</th>
                    <th style="width: 90px;">活跃数量</th>
                    <th style="width: 100px;">活跃率</th>
                    <th style="width: 120px;">客资缺口总额</th>
                    <th style="width: 140px;">预警状态</th>
                </tr>
            </thead>
            <tbody>
        """

        for _, row in disp_sup.iterrows():
            is_warn = row['是否后10名预警']
            row_class = "sup-warning-row" if is_warn else ""
            
            rank = row['排名']
            if rank == 1:
                rank_badge = '<span class="badge-rank1">🥇 第1名</span>'
            elif rank == 2:
                rank_badge = '<span class="badge-rank1" style="background:#f0f5ff; border-color:#d6e4ff; color:#1d39c4;">🥈 第2名</span>'
            elif rank == 3:
                rank_badge = '<span class="badge-rank1" style="background:#fff2e8; border-color:#ffd8bf; color:#d4380d;">🥉 第3名</span>'
            else:
                rank_badge = f"第 {rank} 名"

            gap = row['客资缺口总额']
            if gap < 0:
                gap_html = f'<span class="gap-neg">{gap:,}</span>'
            elif gap > 0:
                gap_html = f'<span class="gap-pos">+{gap:,}</span>'
            else:
                gap_html = '<span>0</span>'

            if is_warn:
                status_badge = '<span class="badge-warning">⚠️ 预警（后10名）</span>'
            else:
                status_badge = '<span class="badge-normal">正常达标</span>'

            sup_html += f"""
                <tr class="{row_class}">
                    <td>{rank_badge}</td>
                    <td><strong>{row['所属督导']}</strong></td>
                    <td>{row['地区']}</td>
                    <td>{row['业务']}</td>
                    <td>{row['门店数量']}</td>
                    <td>{row['活跃数量']}</td>
                    <td><strong style="color: {'#cf1322' if is_warn else '#1f1f1f'};">{row['活跃率']}</strong></td>
                    <td>{gap_html}</td>
                    <td>{status_badge}</td>
                </tr>
            """

        sup_html += """
            </tbody>
        </table>
        """

        st.markdown(sup_html, unsafe_allow_html=True)

if __name__ == "__main__":
    main()



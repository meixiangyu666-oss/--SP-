import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import uuid
import os
import io

# 设置页面配置
st.set_page_config(page_title="SP-批量模版生成工具", page_icon="📊", layout="centered")

# 自定义 CSS 样式
st.markdown("""
    <style>
    /* 主标题样式 */
    .main-title {
        font-size: 2.5em;
        font-weight: bold;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 20px;
    }
    /* 提示文字样式 */
    .instruction {
        font-size: 1.1em;
        color: #34495E;
        margin-bottom: 20px;
    }
    /* 按钮样式 */
    .stButton>button {
        background-color: #3498DB;
        color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 1em;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2980B9;
    }
    /* 下拉菜单样式 */
    .stSelectbox label {
        font-size: 1.1em;
        color: #2C3E50;
        font-weight: bold;
    }
    /* 文件上传框样式 */
    .stFileUploader label {
        font-size: 1.1em;
        color: #2C3E50;
        font-weight: bold;
    }
    /* 成功和错误消息样式 */
    .stSuccess {
        background-color: #E8F5E9;
        border-left: 5px solid #4CAF50;
        padding: 10px;
        border-radius: 5px;
    }
    .stError {
        background-color: #FFEBEE;
        border-left: 5px solid #F44336;
        padding: 10px;
        border-radius: 5px;
    }
    .stWarning {
        background-color: #FFF3E0;
        border-left: 5px solid #FF9800;
        padding: 10px;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# C US 逻辑：从 script-C US.py 提取并调整
def generate_header_from_survey_C(uploaded_file, output_file, country, sheet_name=0):
    try:
        df_survey_C = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.write(f"成功读取文件，数据形状：{df_survey_C.shape}")
        st.write(f"列名列表: {list(df_survey_C.columns)}")
    except FileNotFoundError:
        st.error(f"错误：无法读取上传的文件。请确保文件格式正确。")
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        return None
    
    unique_campaigns = [name for name in df_survey_C['广告活动名称'].dropna() if str(name).strip()]
    st.write(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    non_empty_campaigns = df_survey_C[
        df_survey_C['广告活动名称'].notna() & 
        (df_survey_C['广告活动名称'] != '')
    ]
    required_cols = ['CPC', 'SKU', '广告组默认竞价', '预算']
    if all(col in non_empty_campaigns.columns for col in required_cols):
        campaign_to_values = non_empty_campaigns.drop_duplicates(
            subset='广告活动名称', keep='first'
        ).set_index('广告活动名称')[required_cols].to_dict('index')
    else:
        campaign_to_values = {}
        st.warning(f"警告：缺少列 {set(required_cols) - set(non_empty_campaigns.columns)}，使用默认值")
    
    st.write(f"生成的字典（有 {len(campaign_to_values)} 个活动）: {campaign_to_values}")
    
    keyword_columns = df_survey_C.columns[7:17]
    st.write(f"关键词列: {list(keyword_columns)}")
    
    duplicates_found = False
    st.markdown("### 检查关键词重复")
    for col in keyword_columns:
        col_index = list(df_survey_C.columns).index(col) + 1
        col_letter = chr(64 + col_index) if col_index <= 26 else f"{chr(64 + (col_index-1)//26)}{chr(64 + (col_index-1)%26 + 1)}"
        kw_list = [kw for kw in df_survey_C[col].dropna() if str(kw).strip()]
        
        if len(kw_list) > len(set(kw_list)):
            duplicates_mask = df_survey_C[col].duplicated(keep=False)
            duplicates_df = df_survey_C[duplicates_mask][[col]].dropna()
            st.warning(f"警告：{col_letter} 列 ({col}) 有重复关键词")
            for _, row in duplicates_df.iterrows():
                kw = str(row[col]).strip()
                count = (df_survey_C[col] == kw).sum()
                if count > 1:
                    st.write(f"  重复词: '{kw}' (出现 {count} 次)")
            duplicates_found = True
    
    if duplicates_found:
        st.error("提示：由于检测到关键词重复，生成已终止。请清理重复关键词后重试。")
        return None
    
    st.write("关键词无重复，继续生成...")
    
    neg_exact = list(dict.fromkeys([kw for kw in df_survey_C.get('否定精准', pd.Series()).dropna() if str(kw).strip()]))
    neg_phrase = list(dict.fromkeys([kw for kw in df_survey_C.get('否定词组', pd.Series()).dropna() if str(kw).strip()]))
    suzhu_extra_neg_exact = list(dict.fromkeys([kw for kw in df_survey_C.get('宿主额外否精准', pd.Series()).dropna() if str(kw).strip()]))
    suzhu_extra_neg_phrase = list(dict.fromkeys([kw for kw in df_survey_C.get('宿主额外否词组', pd.Series()).dropna() if str(kw).strip()]))
    
    neg_duplicates_found = False
    st.markdown("### 检查否定关键词重复")
    
    if len(neg_exact) > len(set(neg_exact)):
        neg_duplicates_found = True
        st.warning("警告：'否定精准' 列有重复关键词")
        neg_exact_series = df_survey_C.get('否定精准', pd.Series()).dropna()
        duplicates_mask = neg_exact_series.duplicated(keep=False)
        duplicates_df = df_survey_C[duplicates_mask].loc[:, '否定精准'].dropna()
        for _, row in duplicates_df.items():
            kw = str(row).strip()
            count = (neg_exact_series == kw).sum()
            if count > 1:
                st.write(f"  重复词: '{kw}' (出现 {count} 次)")
    
    if len(neg_phrase) > len(set(neg_phrase)):
        neg_duplicates_found = True
        st.warning("警告：'否定词组' 列有重复关键词")
        neg_phrase_series = df_survey_C.get('否定词组', pd.Series()).dropna()
        duplicates_mask = neg_phrase_series.duplicated(keep=False)
        duplicates_df = df_survey_C[duplicates_mask].loc[:, '否定词组'].dropna()
        for _, row in duplicates_df.items():
            kw = str(row).strip()
            count = (neg_phrase_series == kw).sum()
            if count > 1:
                st.write(f"  重复词: '{kw}' (出现 {count} 次)")
    
    if len(suzhu_extra_neg_exact) > len(set(suzhu_extra_neg_exact)):
        neg_duplicates_found = True
        st.warning("警告：'宿主额外否精准' 列有重复关键词")
        suzhu_exact_series = df_survey_C.get('宿主额外否精准', pd.Series()).dropna()
        duplicates_mask = suzhu_exact_series.duplicated(keep=False)
        duplicates_df = df_survey_C[duplicates_mask].loc[:, '宿主额外否精准'].dropna()
        for _, row in duplicates_df.items():
            kw = str(row).strip()
            count = (suzhu_exact_series == kw).sum()
            if count > 1:
                st.write(f"  重复词: '{kw}' (出现 {count} 次)")
    
    if len(suzhu_extra_neg_phrase) > len(set(suzhu_extra_neg_phrase)):
        neg_duplicates_found = True
        st.warning("警告：'宿主额外否词组' 列有重复关键词")
        suzhu_phrase_series = df_survey_C.get('宿主额外否词组', pd.Series()).dropna()
        duplicates_mask = suzhu_phrase_series.duplicated(keep=False)
        duplicates_df = df_survey_C[duplicates_mask].loc[:, '宿主额外否词组'].dropna()
        for _, row in duplicates_df.items():
            kw = str(row).strip()
            count = (suzhu_phrase_series == kw).sum()
            if count > 1:
                st.write(f"  重复词: '{kw}' (出现 {count} 次)")
    
    if neg_duplicates_found:
        st.error("提示：由于检测到否定关键词重复，生成已终止。请清理重复关键词后重试。")
        return None
    
    st.write("否定关键词无重复，继续生成...")
    
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告组合编号', '广告编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '广告组名称', '开始日期', '结束日期', '投放类型', '状态', '每日预算', 'SKU', '广告组默认竞价',
        '竞价', '关键词文本', '匹配类型', '竞价方案', '广告位', '百分比', '拓展商品投放编号'
    ]
    
    product = '商品推广'
    operation = 'Create'
    status = '已启用'
    targeting_type = '手动'
    bidding_strategy = '动态竞价 - 仅降低'
    default_daily_budget = 12
    default_group_bid = 0.6
    
    keyword_categories = set()
    for col in keyword_columns:
        col_lower = str(col).lower()
        if '/' in col:
            parts = col_lower.split('/')
            if parts[0]:
                keyword_categories.add(parts[0])
            if len(parts) > 1 and parts[1]:
                chinese_part = parts[1].split('-')[0] if '-' in parts[1] else parts[1]
                keyword_categories.add(chinese_part)
        else:
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)]
                    if prefix:
                        keyword_categories.add(prefix)
                        break
    
    keyword_categories.update(['suzhu', '宿主', 'case', '包', 'tape'])
    st.write(f"识别到的关键词类别: {keyword_categories}")
    
    rows = []
    
    for campaign_name in unique_campaigns:
        if campaign_name in campaign_to_values:
            cpc = campaign_to_values[campaign_name]['CPC']
            sku = campaign_to_values[campaign_name]['SKU']
            group_bid = campaign_to_values[campaign_name]['广告组默认竞价']
            budget = campaign_to_values[campaign_name]['预算']
        else:
            cpc = 0.5
            sku = 'SKU-1'
            group_bid = default_group_bid
            budget = default_daily_budget
        
        st.write(f"处理活动: {campaign_name}")
        
        campaign_name_normalized = str(campaign_name).lower()
        matched_category = None
        for category in keyword_categories:
            if category in campaign_name_normalized:
                matched_category = category
                break
        st.write(f"  匹配的关键词类别: {matched_category}")
        
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        is_asin = 'asin' in campaign_name_normalized
        match_type = '精准' if is_exact else '广泛' if is_broad else 'ASIN' if is_asin else None
        st.write(f"  is_exact: {is_exact}, is_broad: {is_broad}, is_asin: {is_asin}, match_type: {match_type}")
        
        keywords = []
        matched_columns = []
        if matched_category and (is_exact or is_broad):
            for col in keyword_columns:
                col_lower = str(col).lower()
                if is_exact and matched_category in col_lower and any(x in col_lower for x in ['精准', 'exact']):
                    matched_columns.append(col)
                    keywords.extend([kw for kw in df_survey_C[col].dropna() if str(kw).strip()])
                elif is_broad and matched_category in col_lower and any(x in col_lower for x in ['广泛', 'broad']):
                    matched_columns.append(col)
                    keywords.extend([kw for kw in df_survey_C[col].dropna() if str(kw).strip()])
            keywords = list(dict.fromkeys(keywords))
            st.write(f"  匹配的列: {matched_columns}")
            st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        else:
            st.write("  无匹配的关键词列，关键词为空")
        
        neg_keywords = []
        if is_broad and matched_category:
            for col in keyword_columns:
                col_lower = str(col).lower()
                if matched_category in col_lower and any(x in col_lower for x in ['精准', 'exact']):
                    neg_keywords.extend([kw for kw in df_survey_C[col].dropna() if str(kw).strip()])
                if any(x in campaign_name_normalized for x in ['suzhu', '宿主']) and any(x in col_lower for x in ['case', '包']) and any(x in col_lower for x in ['精准', 'exact']):
                    neg_keywords.extend([kw for kw in df_survey_C[col].dropna() if str(kw).strip()])
            neg_keywords = list(dict.fromkeys(neg_keywords))
            st.write(f"  精准否定关键词数量: {len(neg_keywords)} (示例: {neg_keywords[:2] if neg_keywords else '无'})")
        
        # 合并 neg_exact 和 neg_keywords，去重
        combined_neg_exact = list(dict.fromkeys(neg_exact + neg_keywords))
        st.write(f"  合并后的否定精准关键词数量: {len(combined_neg_exact)} (示例: {combined_neg_exact[:2] if combined_neg_exact else '无'})")
        
        asin_targets = []
        if is_asin and matched_category:
            potential_asin_cols = []
            for col in df_survey_C.columns:
                col_lower = str(col).lower()
                if matched_category in col_lower and 'asin' in col_lower:
                    potential_asin_cols.append(col)
            
            st.write(f"  潜在ASIN列: {potential_asin_cols}")
            
            if potential_asin_cols:
                def calculate_match_score(col_name, campaign_norm):
                    col_lower = str(col_name).lower()
                    words = re.split(r'[\s/:-]+', col_lower)
                    unique_words = [w.strip() for w in words if w.strip() and w not in ['asin', '精准', '广泛', 'exact', 'broad']]
                    score = sum(1 for word in unique_words if word in campaign_norm)
                    return score, unique_words
                
                scores = {}
                for col in potential_asin_cols:
                    score, words = calculate_match_score(col, campaign_name_normalized)
                    scores[col] = score
                    st.write(f"    列 '{col}' 独特词: {words}, 分数: {score}")
                
                best_col = max(scores, key=scores.get)
                best_score = scores[best_col]
                st.write(f"  选择最佳列: {best_col} (分数: {best_score})")
                
                asin_targets.extend([kw for kw in df_survey_C[best_col].dropna() if str(kw).strip()])
            
            asin_targets = list(dict.fromkeys(asin_targets))
            st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
        
        rows.append([
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', '', targeting_type, status, budget, '', '',
            '', '', '', bidding_strategy, '', '', ''
        ])
        
        rows.append([
            product, '广告组', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', '', group_bid,
            '', '', '', '', '', '', ''
        ])
        
        rows.append([
            product, '商品广告', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', sku, '',
            '', '', '', '', '', '', ''
        ])
        
        if is_exact or is_broad:
            for kw in keywords:
                rows.append([
                    product, '关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, kw, match_type, '', '', '', ''
                ])
        
        if is_broad:
            for kw in combined_neg_exact:  # 使用合并后的否定精准关键词
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            for kw in neg_phrase:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定词组', '', '', '', ''
                ])
            if any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
                for kw in suzhu_extra_neg_exact:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定精准匹配', '', '', '', ''
                    ])
                for kw in suzhu_extra_neg_phrase:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定词组', '', '', '', ''
                    ])
        
        if is_asin:
            for asin in asin_targets:
                rows.append([
                    product, '商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, '', '', '', '', '', f'asin="{asin}"'
                ])
    
    # 创建 DataFrame
    df_header = pd.DataFrame(rows, columns=columns)
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        st.success(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except Exception as e:
        st.error(f"写入文件 {output_file} 时出错：{e}")
        return None

    keyword_rows = [row for row in rows if row[1] == '关键词']
    st.write(f"关键词行数量: {len(keyword_rows)}")
    if keyword_rows:
        st.write(f"示例关键词行: 实体层级={keyword_rows[0][1]}, 关键词文本={keyword_rows[0][19]}, 匹配类型={keyword_rows[0][20]}")
    
    product_targeting_rows = [row for row in rows if row[1] == '商品定向']
    st.write(f"商品定向行数量: {len(product_targeting_rows)}")
    if product_targeting_rows:
        st.write(f"示例商品定向行: 实体层级={product_targeting_rows[0][1]}, 竞价={product_targeting_rows[0][18]}, 拓展商品投放编号={product_targeting_rows[0][24]}")
    
    levels = set(row[1] for row in rows)
    st.write(f"所有实体层级: {levels}")

# B US 逻辑：从 script-B US.py 提取并调整
def generate_header_from_survey_B(uploaded_file, output_file, country, sheet_name=0):
    try:
        # 读取 Excel 文件
        df_survey = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.write(f"成功读取文件，数据形状：{df_survey.shape}")
        st.write(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        st.error(f"错误：无法读取上传的文件。请确保文件格式正确。")
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        return None
    
    # 提取独特活动名称
    unique_campaigns = [name for name in df_survey['广告活动名称'].dropna() if str(name).strip()]
    st.write(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    # 创建活动到 CPC/SKU/广告组默认竞价/预算 的映射
    non_empty_campaigns = df_survey[
        df_survey['广告活动名称'].notna() & 
        (df_survey['广告活动名称'] != '')
    ]
    required_cols = ['CPC', 'SKU', '广告组默认竞价', '预算']
    if all(col in non_empty_campaigns.columns for col in required_cols):
        campaign_to_values = non_empty_campaigns.drop_duplicates(
            subset='广告活动名称', keep='first'
        ).set_index('广告活动名称')[required_cols].to_dict('index')
    else:
        campaign_to_values = {}
        st.warning(f"警告：缺少列 {set(required_cols) - set(non_empty_campaigns.columns)}，使用默认值")
    
    st.write(f"生成的字典（有 {len(campaign_to_values)} 个活动）: {campaign_to_values}")
    
    # 关键词列：第 H 列（索引 7）到第 Q 列（索引 16）
    keyword_columns = df_survey.columns[7:17]
    st.write(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    st.markdown("### 检查关键词重复")
    for col in keyword_columns:
        col_index = list(df_survey.columns).index(col) + 1
        col_letter = chr(64 + col_index) if col_index <= 26 else f"{chr(64 + (col_index-1)//26)}{chr(64 + (col_index-1)%26 + 1)}"
        kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
        
        if len(kw_list) > len(set(kw_list)):
            duplicates_mask = df_survey[col].duplicated(keep=False)
            duplicates_df = df_survey[duplicates_mask][[col]].dropna()
            st.warning(f"警告：{col_letter} 列 ({col}) 有重复关键词")
            for _, row in duplicates_df.iterrows():
                kw = str(row[col]).strip()
                count = (df_survey[col] == kw).sum()
                if count > 1:
                    st.write(f"  重复词: '{kw}' (出现 {count} 次)")
            duplicates_found = True
    
    if duplicates_found:
        st.error("提示：由于检测到关键词重复，生成已终止。请清理重复关键词后重试。")
        return None
    
    st.write("关键词无重复，继续生成...")
    
    # 否定关键词聚合（去重）
    neg_exact = list(dict.fromkeys([kw for kw in df_survey.get('否定精准', pd.Series()).dropna() if str(kw).strip()]))
    neg_phrase = list(dict.fromkeys([kw for kw in df_survey.get('否定词组', pd.Series()).dropna() if str(kw).strip()]))
    suzhu_extra_neg_exact = list(dict.fromkeys([kw for kw in df_survey.get('宿主额外否精准', pd.Series()).dropna() if str(kw).strip()]))
    suzhu_extra_neg_phrase = list(dict.fromkeys([kw for kw in df_survey.get('宿主额外否词组', pd.Series()).dropna() if str(kw).strip()]))
    neg_asin = list(dict.fromkeys([kw for kw in df_survey.get('否定ASIN', pd.Series()).dropna() if str(kw).strip()]))
    
    # 定义关键词类别到精准词列的映射
    keyword_categories = {
        'suzhu': 'suzhu/宿主-精准词',
        '宿主': 'suzhu/宿主-精准词',
        'case': 'case/包-精准词',
        '包': 'case/包-精准词',
        'cards': 'cards精准词',
        'acces': 'acces精准词',
        'acc': 'acc精准词',
        None: '精准词'  # XX 组，默认列
    }
    
    # 提取精准关键词
    exact_keywords = {key: list(dict.fromkeys([kw for kw in df_survey.get(col, pd.Series()).dropna() if str(kw).strip()]))
                      for key, col in keyword_categories.items() if col in df_survey.columns}
    
    # 列定义
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告组合编号', '广告编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '广告组名称', '开始日期', '结束日期', '投放类型', '状态', '每日预算', 'SKU', '广告组默认竞价',
        '竞价', '关键词文本', '匹配类型', '竞价方案', '广告位', '百分比', '拓展商品投放编号'
    ]
    
    # 默认值
    product = '商品推广'
    operation = 'Create'
    status = '已启用'
    targeting_type = '手动'
    bidding_strategy = '动态竞价 - 仅降低'
    default_daily_budget = 12
    default_group_bid = 0.6
    
    # 初始化结果列表
    rows = []
    
    # 处理每个广告活动
    for campaign_name in unique_campaigns:
        campaign_name_normalized = str(campaign_name).lower()
        
        # 获取 CPC、SKU、预算和广告组默认竞价
        campaign_values = campaign_to_values.get(campaign_name, {})
        cpc = campaign_values.get('CPC', 0.6)
        sku = campaign_values.get('SKU', '')
        daily_budget = campaign_values.get('预算', default_daily_budget)
        group_bid = campaign_values.get('广告组默认竞价', default_group_bid)
        
        # 判断匹配类型
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        is_asin = 'asin' in campaign_name_normalized
        match_type = '精准' if is_exact else '广泛' if is_broad else 'ASIN' if is_asin else None
        st.write(f"处理活动: {campaign_name}")
        st.write(f"  is_exact: {is_exact}, is_broad: {is_broad}, is_asin: {is_asin}, match_type: {match_type}")
        
        # 提取关键词类别
        keyword_categories_set = set(keyword_categories.keys()) - {None}
        st.write(f"识别到的关键词类别: {keyword_categories_set}")
        
        # 匹配关键词类别 - 去掉动态匹配，改为匹配到配件组
        matched_category = None
        matched_columns = []

        # 首先尝试预定义的映射
        for category in keyword_categories_set:
            if category in campaign_name_normalized:
                matched_category = category
                # 根据匹配类型找到对应的列
                if is_exact:
                    target_col = keyword_categories[category]
                    if target_col in df_survey.columns:
                        matched_columns.append(target_col)
                elif is_broad:
                    # 查找对应的广泛词列
                    target_col_broad = keyword_categories[category].replace('精准', '广泛')
                    if target_col_broad in df_survey.columns:
                        matched_columns.append(target_col_broad)
                break

        # 如果没有匹配到预定义组别，则匹配到配件组
        if not matched_columns and (is_exact or is_broad):
            matched_category = '配件'
            if is_exact:
                # 配件精准组：使用 L 列（索引11）
                target_col = df_survey.columns[11]  # L列
                if target_col in df_survey.columns:
                    matched_columns.append(target_col)
                    st.write(f"  匹配到配件精准组，使用列: {target_col}")
            elif is_broad:
                # 配件广泛组：使用 M 列（索引12）
                target_col = df_survey.columns[12]  # M列
                if target_col in df_survey.columns:
                    matched_columns.append(target_col)
                    st.write(f"  匹配到配件广泛组，使用列: {target_col}")

        st.write(f"  匹配的关键词类别: {matched_category}")
        
        # 提取关键词
        keywords = []
        if matched_columns and (is_exact or is_broad):
            for col in matched_columns:
                if col in df_survey.columns:
                    col_keywords = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
                    keywords.extend(col_keywords)
                    st.write(f"  从列 {col} 提取 {len(col_keywords)} 个关键词")
            
            keywords = list(dict.fromkeys(keywords))  # 去重
            st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        else:
            st.write("  无匹配的关键词列，关键词为空")
        
        # 初始化否定关键词列表
        campaign_neg_exact = []
        campaign_neg_phrase = []
        
        # 根据组别和匹配类型设置否定关键词
        if is_exact and any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
            # 宿主精准组：仅通用否定精准
            campaign_neg_exact = list(dict.fromkeys(neg_exact))
            campaign_neg_phrase = list(dict.fromkeys(neg_phrase))
        elif is_exact:
            # 其他精准组：通用否定精准 + 通用否定词组
            campaign_neg_exact = list(dict.fromkeys(neg_exact))
            campaign_neg_phrase = list(dict.fromkeys(neg_phrase))
        elif is_broad:
            # 广泛组：通用否定精准 + 通用否定词组 + 对应精准组关键词（作为否定精准）
            campaign_neg_exact = list(dict.fromkeys(neg_exact))
            campaign_neg_phrase = list(dict.fromkeys(neg_phrase))
            
            # 为广泛组添加对应精准组的否定关键词
            if matched_category in exact_keywords and matched_category != '配件':
                # 预定义广泛组：添加对应精准组关键词
                exact_kws = exact_keywords.get(matched_category, [])
                campaign_neg_exact.extend(exact_kws)
                st.write(f"  为预定义广泛组添加 {len(exact_kws)} 个 {matched_category} 精准词作为否定精准词")
            elif matched_category == '配件':
                # 配件广泛组：添加配件精准组关键词
                accessory_exact_col = df_survey.columns[11]  # L列
                if accessory_exact_col in df_survey.columns:
                    accessory_exact_kws = list(dict.fromkeys([kw for kw in df_survey[accessory_exact_col].dropna() if str(kw).strip()]))
                    campaign_neg_exact.extend(accessory_exact_kws)
                    st.write(f"  为配件广泛组添加 {len(accessory_exact_kws)} 个配件精准词作为否定精准词 (从列: {accessory_exact_col})")
            
            campaign_neg_exact = list(dict.fromkeys(campaign_neg_exact))  # 去重
        
        # 为宿主组添加额外否定关键词（如果不是宿主精准组）
        if not (is_exact and any(x in campaign_name_normalized for x in ['suzhu', '宿主'])):
            if any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
                campaign_neg_exact.extend(suzhu_extra_neg_exact)
                campaign_neg_phrase.extend(suzhu_extra_neg_phrase)
                campaign_neg_exact = list(dict.fromkeys(campaign_neg_exact))
                campaign_neg_phrase = list(dict.fromkeys(campaign_neg_phrase))
        
        # 生成广告活动行
        rows.append([
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', '', targeting_type, status, daily_budget, '',
            '', '', '', '', bidding_strategy, '', '', ''
        ])
        
        # 生成广告组行
        rows.append([
            product, '广告组', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', '', group_bid,
            '', '', '', '', '', '', ''
        ])
        
        # 生成商品广告行
        rows.append([
            product, '商品广告', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', sku, '',
            '', '', '', '', '', '', ''
        ])
        
        # 生成关键词行
        if is_exact or is_broad:
            for kw in keywords:
                rows.append([
                    product, '关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, kw, match_type, '', '', '', ''
                ])
        
        # 生成否定关键词行
        if is_exact or is_broad:
            # 精准组和广泛组：添加否定精准和否定词组
            for kw in campaign_neg_exact:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            for kw in campaign_neg_phrase:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定词组', '', '', '', ''
                ])
        
        # 生成商品定向和否定商品定向（仅 ASIN 组）
        if is_asin:
            asin_targets = []
            # 精确匹配：列名必须与广告活动名称完全一致
            if campaign_name in df_survey.columns:
                asin_targets.extend([asin for asin in df_survey[campaign_name].dropna() if str(asin).strip()])
                st.write(f"  找到与活动名称完全匹配的列: {campaign_name}")
            else:
                st.write(f"  未找到与活动名称完全匹配的列: {campaign_name}")
                
            asin_targets = list(dict.fromkeys(asin_targets))
            st.write(f"  ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
            
            for asin in asin_targets:
                rows.append([
                    product, '商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, '', '', '', '', '', f'asin="{asin}"'
                ])
            for asin in neg_asin:
                rows.append([
                    product, '否定商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    '', '', '', '', '', '', f'asin="{asin}"'
                ])
    
    # 创建 DataFrame
    df_header = pd.DataFrame(rows, columns=columns)
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        st.success(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except Exception as e:
        st.error(f"写入文件 {output_file} 时出错：{e}")
        return None
    
    # 调试输出
    keyword_rows = [row for row in rows if row[1] == '关键词']
    st.write(f"关键词行数量: {len(keyword_rows)}")
    if keyword_rows:
        st.write(f"示例关键词行: 实体层级={keyword_rows[0][1]}, 关键词文本={keyword_rows[0][19]}, 匹配类型={keyword_rows[0][20]}")
    
    product_targeting_rows = [row for row in rows if row[1] == '商品定向']
    st.write(f"商品定向行数量: {len(product_targeting_rows)}")
    if product_targeting_rows:
        st.write(f"示例商品定向行: 实体层级={product_targeting_rows[0][1]}, 竞价={product_targeting_rows[0][18]}, 拓展商品投放编号={product_targeting_rows[0][24]}")
    
    levels = set(row[1] for row in rows)
    st.write(f"所有实体层级: {levels}")

# Streamlit 界面
st.markdown('<div class="main-title">SP-批量模版生成工具</div>', unsafe_allow_html=True)
st.markdown('<div class="instruction">请选择国家并上传 Excel 文件，点击按钮生成对应的 Header 文件（支持任意文件名的 .xlsx 文件）。<br>Please select a country and upload an Excel file, then click the button to generate the corresponding Header file (supports any .xlsx filename).</div>', unsafe_allow_html=True)

# 国家选择
country = st.selectbox("选择国家 / Select Country", ["C US", "B US"])

# 文件上传
uploaded_file = st.file_uploader("上传 Excel 文件 / Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # 动态生成输出文件名
    output_file = f"header-{country.replace(' ', '_')}.xlsx"
    
    # 运行按钮
    if st.button("生成 Header 文件 / Generate Header File"):
        with st.spinner("正在处理文件... / Processing file..."):
            if country == "C US":
                result = generate_header_from_survey_C(uploaded_file, output_file, country)
            elif country == "B US":
                result = generate_header_from_survey_B(uploaded_file, output_file, country)
            else:
                st.error("不支持的国家选择。")
                result = None
            
            if result and os.path.exists(result):
                with open(result, "rb") as f:
                    st.download_button(
                        label=f"下载 {output_file} / Download {output_file}",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                # 调试信息
                st.markdown("### 处理结果 / Processing Results")
                df_result = pd.read_excel(result)
                keyword_rows = [row for row in df_result.to_dict('records') if row['实体层级'] == '关键词']
                st.write(f"关键词行数量 / Keyword Rows: {len(keyword_rows)}")
                if keyword_rows:
                    st.write(f"示例关键词行 / Example Keyword Row: 实体层级={keyword_rows[0]['实体层级']}, 关键词文本={keyword_rows[0]['关键词文本']}, 匹配类型={keyword_rows[0]['匹配类型']}")
                product_targeting_rows = [row for row in df_result.to_dict('records') if row['实体层级'] == '商品定向']
                st.write(f"商品定向行数量 / Product Targeting Rows: {len(product_targeting_rows)}")
                if product_targeting_rows:
                    st.write(f"示例商品定向行 / Example Product Targeting Row: 实体层级={product_targeting_rows[0]['实体层级']}, 竞价={product_targeting_rows[0]['竞价']}, 拓展商品投放编号={product_targeting_rows[0]['拓展商品投放编号']}")
                levels = set(row['实体层级'] for row in df_result.to_dict('records'))
                st.write(f"所有实体层级 / All Entity Levels: {levels}")
            else:
                st.error("生成文件失败，请检查上传的文件格式或内容。 / Failed to generate file, please check the file format or content.")
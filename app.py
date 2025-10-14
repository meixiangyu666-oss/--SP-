import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import uuid
import os

# 通用函数：从调研 Excel 生成表头 Excel
def generate_header_from_survey(uploaded_file, output_file, country, sheet_name=0):
    try:
        # 读取上传的 Excel 文件
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
    st.write("\n=== 检查关键词重复 ===")
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
        st.error("提示：由于检测到关键词重复，本次不生成表格。请清理重复后重试。")
        return None
    
    st.write("关键词无重复，继续生成...")
    
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
    
    # 提取关键词类别（JP 和 K EU 通用逻辑）
    def extract_keyword_categories(df_survey):
        categories = set()
        for col in df_survey.columns:
            col_lower = str(col).lower()
            if any(x in col_lower for x in ['精准词', '广泛词', '精准', '广泛']):
                for suffix in ['精准词', '广泛词', '精准', '广泛']:
                    if col_lower.endswith(suffix):
                        prefix = col_lower[:-len(suffix)].strip()
                        parts = re.split(r'[/\-_\s\.]', prefix)
                        for part in parts:
                            if part and len(part) > 1:
                                categories.add(part)
                        break
            elif 'asin' in col_lower and '否定' not in col_lower:
                prefix = col_lower.replace('asin', '').strip()
                parts = re.split(r'[/\-_\s\.]', prefix)
                for part in parts:
                    if part and len(part) > 1:
                        categories.add(part)
        categories.update(['suzhu', 'host', '宿主', 'case', '包', 'tape'])
        categories.discard('')
        return categories
    
    keyword_categories = extract_keyword_categories(df_survey)
    st.write(f"识别到的关键词类别: {keyword_categories}")
    
    # JP 特定逻辑：否定关键词
    def get_jp_neg_keywords(df_survey):
        neg_exact = [kw for kw in df_survey.get('否定精准', pd.Series()).dropna() if str(kw).strip()]
        neg_phrase = [kw for kw in df_survey.get('否定词组', pd.Series()).dropna() if str(kw).strip()]
        suzhu_extra_neg_exact = [kw for kw in df_survey.get('宿主额外否精准', pd.Series()).dropna() if str(kw).strip()]
        suzhu_extra_neg_phrase = [kw for kw in df_survey.get('宿主额外否词组', pd.Series()).dropna() if str(kw).strip()]
        neg_asin = [kw for kw in df_survey.get('否定ASIN', pd.Series()).dropna() if str(kw).strip()]
        return neg_exact, neg_phrase, suzhu_extra_neg_exact, suzhu_extra_neg_phrase, neg_asin
    
    # K EU 特定逻辑：否定关键词
    def get_k_eu_neg_keywords(df_survey, campaign_name, matched_category, is_broad, is_exact):
        neg_exact = []
        neg_phrase = []
        if is_broad:
            s_col = df_survey.iloc[:, 18]  # S列
            t_col = df_survey.iloc[:, 19]  # T列
            neg_exact = [kw for kw in s_col.dropna() if str(kw).strip()]
            neg_phrase = [kw for kw in t_col.dropna() if str(kw).strip()]
        elif is_exact and matched_category:
            if matched_category in ['suzhu', 'host', '宿主']:
                u_col = df_survey.iloc[:, 20]  # U列
                v_col = df_survey.iloc[:, 21]  # V列
                neg_exact = [kw for kw in u_col.dropna() if str(kw).strip()]
                neg_phrase = [kw for kw in v_col.dropna() if str(kw).strip()]
            elif matched_category == 'case':
                w_col = df_survey.iloc[:, 22]  # W列
                x_col = df_survey.iloc[:, 23]  # X列
                neg_exact = [kw for kw in w_col.dropna() if str(kw).strip()]
                neg_phrase = [kw for kw in x_col.dropna() if str(kw).strip()]
        neg_exact = list(dict.fromkeys(neg_exact))
        neg_phrase = list(dict.fromkeys(neg_phrase))
        neg_asin = [kw for kw in df_survey.get('否定ASIN', pd.Series()).dropna() if str(kw).strip()]
        st.write(f"否定关键词：精准 {len(neg_exact)} 个，词组 {len(neg_phrase)} 个，否定ASIN {len(neg_asin)} 个")
        return neg_exact, neg_phrase, neg_asin
    
    # 函数：查找匹配的ASIN列（K EU 逻辑，包含颜色匹配）
    def find_matching_asin_columns_k_eu(campaign_name, df_survey, keyword_categories):
        campaign_name_normalized = str(campaign_name).lower()
        if 'asin' not in campaign_name_normalized:
            st.write(f"  {campaign_name} 不是商品定向活动，无匹配ASIN列")
            return []
        
        sorted_categories = sorted(keyword_categories, key=len)
        matched_category = None
        for category in sorted_categories:
            if category in campaign_name_normalized:
                matched_category = category
                break
        
        if not matched_category:
            color_words = ['红', '白', '黑', '蓝']
            for color in color_words:
                if color in campaign_name_normalized:
                    matched_category = color
                    st.write(f"  Fallback 匹配颜色类别: {matched_category}")
                    break
        
        if not matched_category:
            st.write(f"  {campaign_name} 未匹配到任何关键词类别，无匹配ASIN列")
            return []
        
        st.write(f"  匹配的关键词类别: {matched_category}")
        
        color = None
        color_words = ['红', '白', '黑', '蓝']
        for c in color_words:
            if c in campaign_name_normalized:
                color = c
                break
        st.write(f"  提取的颜色: {color}")
        
        words = re.findall(r'[a-zA-Z0-9\u4e00-\u9fff]+', campaign_name_normalized)
        exclude_words = {matched_category, 'asin', '商品定向', '定向', '精准', '广泛', 'exact', 'broad', 'host', 'case'} if matched_category else {'asin', '商品定向', '定向', '精准', '广泛', 'exact', 'broad', 'host', 'case'}
        candidate_words = [word for word in words if word not in exclude_words and len(word) > 1]
        st.write(f"  候选匹配词: {candidate_words}")
        
        matching_columns = []
        for col in df_survey.columns:
            col_lower = str(col).lower()
            if (matched_category in col_lower and 
                'asin' in col_lower and 
                '否定' not in col_lower and
                (not color or color in col_lower)):
                matching_columns.append(col)
        
        st.write(f"  初步匹配的ASIN列: {matching_columns}")
        
        if len(matching_columns) > 1:
            best_match = None
            max_matches = 0
            for col in matching_columns:
                col_lower = str(col).lower()
                match_count = sum(1 for word in candidate_words if word in col_lower)
                if match_count > max_matches:
                    max_matches = match_count
                    best_match = col
                elif match_count == max_matches and best_match:
                    best_match = col if len(col_lower) > len(best_match.lower()) else best_match
            if best_match:
                matching_columns = [best_match]
                st.write(f"  精细匹配后选择列: {matching_columns}")
            else:
                st.write(f"  无法进一步筛选，保留初步匹配列: {matching_columns}")
        
        return matching_columns
    
    # 函数：查找匹配的关键词列（JP 和 K EU 通用）
    def find_matching_keyword_columns(campaign_name, df_survey, keyword_categories, keyword_columns, match_type):
        campaign_name_normalized = str(campaign_name).lower()
        matched_categories = []
        for category in keyword_categories:
            if category and category in campaign_name_normalized:
                matched_categories.append(category)
        
        st.write(f"  匹配的关键词类别: {matched_categories}")
        
        if not matched_categories:
            st.write("  无匹配的关键词类别")
            return [], []
        
        match_type_keywords = ['精准', 'exact'] if match_type == '精准' else ['广泛', 'broad']
        matching_columns = []
        for col in keyword_columns:
            col_lower = str(col).lower()
            has_match_type = any(keyword in col_lower for keyword in match_type_keywords)
            has_category = any(category in col_lower for category in matched_categories)
            if has_match_type and has_category:
                matching_columns.append(col)
        
        st.write(f"  匹配的列: {matching_columns}")
        
        keywords = []
        for col in matching_columns:
            keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        keywords = list(dict.fromkeys(keywords))
        st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        
        return matching_columns, keywords
    
    # 函数：查找交叉否定关键词（JP 逻辑）
    def find_cross_neg_keywords_jp(campaign_name, df_survey, keyword_categories, keyword_columns):
        campaign_name_normalized = str(campaign_name).lower()
        cross_neg_keywords = []
        if any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
            for col in keyword_columns:
                col_lower = str(col).lower()
                if any(case_word in col_lower for case_word in ['case', '包', 'tape']) and any(x in col_lower for x in ['精准', 'exact']):
                    cross_neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        elif any(x in campaign_name_normalized for x in ['case', '包', 'tape']):
            for col in keyword_columns:
                col_lower = str(col).lower()
                if any(suzhu_word in col_lower for suzhu_word in ['suzhu', '宿主']) and any(x in col_lower for x in ['精准', 'exact']):
                    cross_neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        cross_neg_keywords = list(dict.fromkeys(cross_neg_keywords))
        st.write(f"  交叉否定关键词数量: {len(cross_neg_keywords)} (示例: {cross_neg_keywords[:2] if cross_neg_keywords else '无'})")
        return cross_neg_keywords
    
    # 函数：查找否定关键词（K EU 逻辑）
    def find_neg_keywords_k_eu(campaign_name, df_survey, keyword_categories, keyword_columns):
        campaign_name_normalized = str(campaign_name).lower()
        sorted_categories = sorted(keyword_categories, key=len)
        matched_category = None
        for category in sorted_categories:
            if category in campaign_name_normalized:
                matched_category = category
                break
        if not matched_category:
            return []
        neg_keywords = []
        for col in keyword_columns:
            col_lower = str(col).lower()
            if matched_category in col_lower and any(x in col_lower for x in ['精准', 'exact']):
                neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        neg_keywords = list(dict.fromkeys(neg_keywords))
        st.write(f"  精准否定关键词数量: {len(neg_keywords)} (示例: {neg_keywords[:2] if neg_keywords else '无'})")
        return neg_keywords
    
    # 生成数据行
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
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        is_asin = 'asin' in campaign_name_normalized
        match_type = '精准' if is_exact else '广泛' if is_broad else 'ASIN' if is_asin else None
        st.write(f"  is_exact: {is_exact}, is_broad: {is_broad}, is_asin: {is_asin}, match_type: {match_type}")
        
        # 提取关键词
        keywords = []
        matched_columns = []
        if is_exact or is_broad:
            matched_columns, keywords = find_matching_keyword_columns(
                campaign_name, df_survey, keyword_categories, keyword_columns, match_type
            )
        
        # 提取否定关键词
        neg_exact = []
        neg_phrase = []
        neg_asin = []
        if country == 'JP':
            neg_exact, neg_phrase, suzhu_extra_neg_exact, suzhu_extra_neg_phrase, neg_asin = get_jp_neg_keywords(df_survey)
        elif country == 'K EU':
            matched_category = next((cat for cat in sorted(keyword_categories, key=len) if cat in campaign_name_normalized), None)
            neg_exact, neg_phrase, neg_asin = get_k_eu_neg_keywords(df_survey, campaign_name, matched_category, is_broad, is_exact)
        
        # 提取 ASIN
        asin_targets = []
        if is_asin:
            matching_columns = find_matching_asin_columns_k_eu(campaign_name, df_survey, keyword_categories) if country == 'K EU' else find_matching_keyword_columns(campaign_name, df_survey, keyword_categories, keyword_columns, 'ASIN')[0]
            for col in matching_columns:
                asin_targets.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
            asin_targets = list(dict.fromkeys(asin_targets))
            st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
        
        # K EU 特有：竞价调整行
        if country == 'K EU':
            placement_value = "广告位：商品页面" if is_asin else "广告位：搜索结果首页首位"
            rows.append([
                product, '竞价调整', operation, campaign_name, '', '', '', '', '',
                campaign_name, campaign_name, '', '', targeting_type, '', '', '', '',
                '', '', '', bidding_strategy, placement_value, '900', ''
            ])
            st.write(f"  添加竞价调整行: 广告位={placement_value}")
        
        # 广告活动行
        rows.append([
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', '', targeting_type, status, budget, '', '',
            '', '', '', bidding_strategy, '', '', ''
        ])
        
        # 广告组行
        rows.append([
            product, '广告组', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', '', group_bid,
            '', '', '', '', '', '', ''
        ])
        
        # 商品广告行
        rows.append([
            product, '商品广告', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', sku, '',
            '', '', '', '', '', '', ''
        ])
        
        # 关键词行
        if is_exact or is_broad:
            for kw in keywords:
                rows.append([
                    product, '关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, kw, match_type, '', '', '', ''
                ])
        
        # 否定关键词行
        if is_exact or is_broad:
            for kw in neg_exact:
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
            
            # JP 特有：交叉否定和宿主额外否定
            if country == 'JP':
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
                cross_neg_keywords = find_cross_neg_keywords_jp(campaign_name, df_survey, keyword_categories, keyword_columns)
                for kw in cross_neg_keywords:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定精准匹配', '', '', '', ''
                    ])
            
            # K EU 特有：广泛组否定精准关键词
            if country == 'K EU' and is_broad:
                neg_keywords = find_neg_keywords_k_eu(campaign_name, df_survey, keyword_categories, keyword_columns)
                for kw in neg_keywords:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定精准匹配', '', '', '', ''
                    ])
        
        # 商品定向和否定商品定向
        if is_asin:
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

# Streamlit 界面
st.title("Excel Header 生成工具")
st.write("请选择国家并上传 Excel 文件，点击按钮生成对应的 Header 文件。")

# 国家选择
country = st.selectbox("选择国家", ["JP", "K EU"])

# 文件上传
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    # 动态生成输出文件名
    output_file = f"header-{country.replace(' ', '_')}.xlsx"
    
    # 运行按钮
    if st.button("生成 Header 文件"):
        with st.spinner("正在处理文件..."):
            result = generate_header_from_survey(uploaded_file, output_file, country)
            if result and os.path.exists(result):
                with open(result, "rb") as f:
                    st.download_button(
                        label=f"下载 {output_file}",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                # 调试信息
                keyword_rows = [row for row in pd.read_excel(result).to_dict('records') if row['实体层级'] == '关键词']
                st.write(f"关键词行数量: {len(keyword_rows)}")
                if keyword_rows:
                    st.write(f"示例关键词行: 实体层级={keyword_rows[0]['实体层级']}, 关键词文本={keyword_rows[0]['关键词文本']}, 匹配类型={keyword_rows[0]['匹配类型']}")
                product_targeting_rows = [row for row in pd.read_excel(result).to_dict('records') if row['实体层级'] == '商品定向']
                st.write(f"商品定向行数量: {len(product_targeting_rows)}")
                if product_targeting_rows:
                    st.write(f"示例商品定向行: 实体层级={product_targeting_rows[0]['实体层级']}, 竞价={product_targeting_rows[0]['竞价']}, 拓展商品投放编号={product_targeting_rows[0]['拓展商品投放编号']}")
                if country == 'K EU':
                    bid_adjustment_rows = [row for row in pd.read_excel(result).to_dict('records') if row['实体层级'] == '竞价调整']
                    st.write(f"竞价调整行数量: {len(bid_adjustment_rows)}")
                    if bid_adjustment_rows:
                        st.write(f"示例竞价调整行: 实体层级={bid_adjustment_rows[0]['实体层级']}, 广告位={bid_adjustment_rows[0]['广告位']}, 百分比={bid_adjustment_rows[0]['百分比']}")
                levels = set(row['实体层级'] for row in pd.read_excel(result).to_dict('records'))
                st.write(f"所有实体层级: {levels}")
            else:
                st.error("生成文件失败，请检查上传的文件格式或内容。")

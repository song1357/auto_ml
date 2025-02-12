import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx2pdf import convert
import io
from sklearn.preprocessing import StandardScaler
from kmodes.kmodes import KModes
from sklearn.cluster import KMeans, MiniBatchKMeans
from sklearn.metrics import silhouette_score
import time
import tempfile
import os
import base64
from docx.oxml.ns import qn
from functools import lru_cache
import pythoncom
from docx.oxml import parse_xml
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
from docx.shared import Cm

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 添加缓存装饰器来优化性能
@st.cache_data
def load_data(file):
    try:
        # 读取所有sheet的名称
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        # 创建进度条
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        # 读取每个sheet的数据
        all_data = {}
        for i, sheet in enumerate(sheet_names):
            progress_text.text(f'正在加载工作表: {sheet}')
            progress_bar.progress((i + 1) / len(sheet_names))
            all_data[sheet] = pd.read_excel(file, sheet_name=sheet)
        
        # 清除进度条和文本
        progress_text.empty()
        progress_bar.empty()
        
        # 显示成功消息
        st.success('数据加载完成！')
        
        return all_data, sheet_names
    except Exception as e:
        st.error(f'加载数据时出错: {str(e)}')
        return None, None

@st.cache_data
def preprocess_data(df, selected_features):
    """数据预处理"""
    # 选择特征
    X = df[selected_features].copy()
    
    # 标准化
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    return pd.DataFrame(X_scaled, columns=X.columns)

@st.cache_data
def get_clustering_model(model_name, n_clusters):
    """获取聚类模型"""
    if model_name == 'KModes':
        return KModes(n_clusters=n_clusters, init='Huang', n_init=5, verbose=0)
    elif model_name == 'MiniBatchKMeans':
        return MiniBatchKMeans(n_clusters=n_clusters, random_state=42)
    else:
        return KMeans(n_clusters=n_clusters, random_state=42)

def find_optimal_k(X, max_k=10, sample_size=3000):
    """
    寻找最优K值，使用采样方法处理大数据集
    """
    K = range(2, max_k + 1)
    silhouette_scores = []
    costs = []
    n_samples = len(X)
    
    # 如果数据集太大，使用采样进行轮廓系数计算
    use_sampling = n_samples > sample_size
    if use_sampling:
        # 随机采样
        indices = np.random.choice(n_samples, sample_size, replace=False)
        X_sample = X.iloc[indices].values
        X_values = X.values
    else:
        X_values = X.values
    
    for k in K:
        # KMeans聚类
        kmeans = KMeans(n_clusters=k, random_state=42)
        kmeans.fit(X_values)
        
        # 计算轮廓系数（使用采样数据）
        if use_sampling:
            sample_labels = kmeans.predict(X_sample)
            silhouette_scores.append(silhouette_score(X_sample, sample_labels))
        else:
            silhouette_scores.append(silhouette_score(X_values, kmeans.labels_))
        
        # 计算组内平方和
        costs.append(kmeans.inertia_)
    
    # 找到最优k值（使用肘部法则）
    costs_diff = np.diff(costs)
    costs_diff_r = np.diff(costs[::-1])
    optimal_k = K[len(costs_diff_r[costs_diff_r > np.mean(costs_diff_r)]) + 1]
    
    return K, silhouette_scores, costs, optimal_k

def create_evaluation_plots(K, costs, silhouette_scores, current_k):
    """创建评估图"""
    fig = plt.figure(figsize=(12, 5))
    
    # 肘部法则图
    ax1 = plt.subplot(121)
    ax1.plot(K, costs, 'bx-')
    ax1.axvline(x=current_k, color='r', linestyle='--', alpha=0.5)
    ax1.set_xlabel('K值 (聚类数量)')
    ax1.set_ylabel('组内平方和')
    ax1.set_title('肘部法则分析')
    
    # 轮廓系数图
    ax2 = plt.subplot(122)
    ax2.plot(K, silhouette_scores, 'rx-')
    ax2.axvline(x=current_k, color='r', linestyle='--', alpha=0.5)
    ax2.set_xlabel('K值 (聚类数量)')
    ax2.set_ylabel('轮廓系数')
    ax2.set_title('轮廓系数分析')
    
    plt.tight_layout()
    return fig

def create_clustering_report(df, filename, cuisine_col_idx, level_col_idx, selected_features, k_range, model_name, output_format='docx', selected_cuisine='全部', selected_level='全部'):
    """创建聚类分析报告"""
    # 创建一个新的文档
    doc = Document()
    
    # 添加标题
    doc.add_heading('调味品使用聚类分析报告', 0)
    
    # 基本信息表格
    doc.add_heading('1. 基本信息', level=1)
    info_table = doc.add_table(rows=1, cols=2)
    info_table.style = 'Table Grid'
    
    # 设置表头
    header_cells = info_table.rows[0].cells
    header_cells[0].text = '项目'
    header_cells[1].text = '内容'
    
    # 设置表头样式
    for cell in header_cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = 1
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
        tcPr.append(tcVAlign)
    
    # 填充基本信息
    info_data = [
        ('数据文件', filename),
        ('选择的特征', ', '.join(selected_features)),
        ('聚类范围', f'K = {k_range[0]} ~ {k_range[-1]}'),
        ('分析时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ('选择的菜系', selected_cuisine),
        ('选择的档次', selected_level)
    ]
    
    for label, value in info_data:
        row = info_table.add_row()
        cells = row.cells
        cells[0].text = label
        cells[1].text = str(value)
        # 设置单元格样式
        for cell in cells:
            cell.paragraphs[0].alignment = 1
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
            tcPr.append(tcVAlign)
    
    # 数据预处理
    X = preprocess_data(df, selected_features)
    
    # 2. 聚类数量评估
    doc.add_heading('2. 聚类数量评估', level=1)
    K, silhouette_scores, costs, optimal_k = find_optimal_k(X, max(k_range))
    
    # 绘制评估图
    fig = create_evaluation_plots(K, costs, silhouette_scores, optimal_k)
    
    # 保存图片
    img_path = os.path.join(tempfile.gettempdir(), 'evaluation.png')
    plt.savefig(img_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    # 添加图片到文档
    paragraph = doc.add_paragraph()
    # 设置段落格式
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format.space_before = Pt(8)
    paragraph_format.space_after = Pt(12)
    # 添加图片
    run = paragraph.add_run()
    run.add_picture(img_path, width=Inches(7))
    os.remove(img_path)
    
    # 添加评估图解释
    doc.add_heading('2.1 评估图解释', level=2)
    
    # 创建评估图解释表格
    table = doc.add_table(rows=2, cols=2)
    table.style = 'Table Grid'
    
    # 设置表格标题
    cells = table.rows[0].cells
    cells[0].text = '肘部法则图解释'
    cells[1].text = '轮廓系数解释'
    
    # 设置表头样式
    for cell in cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = 1
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
        tcPr.append(tcVAlign)
    
    # 添加解释内容
    elbow_content = """1. 肘部法则通过观察组内平方和（WCSS）的变化来评估聚类效果。
2. WCSS表示组内样本与其质心之间的距离平方和，值越小表示聚类越紧密。
3. 当增加聚类数量时，WCSS会持续下降。
4. 在某个K值处，WCSS的下降速率会显著减缓，形成"肘部"。
5. 这个"肘部"位置通常被认为是较好的聚类数量选择。"""

    silhouette_content = """1. 轮廓系数评估聚类的内聚度和分离度，取值范围为[-1, 1]。
2. 值越接近1，表示样本与自己所在的组更相似，与其他组更不相似。
3. 值越接近0，表示样本位于组的边界。
4. 值越接近-1，表示样本可能被分配到了错误的组。
5. 一般认为轮廓系数大于0.5表示聚类效果较好。"""

    cells = table.rows[1].cells
    cells[0].text = elbow_content
    cells[1].text = silhouette_content
    
    # 设置内容单元格样式
    for cell in cells:
        cell.paragraphs[0].alignment = 1
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
        tcPr.append(tcVAlign)
    
    # 3. 不同K值的聚类分析
    doc.add_heading('3. 不同K值的聚类分析', level=1)
    
    # 获取列名
    cuisine_col = df.columns[cuisine_col_idx]
    level_col = df.columns[level_col_idx]
    
    for k in k_range:
        doc.add_heading(f'K={k}的聚类分析', level=2)
        
        # 执行聚类
        model = get_clustering_model(model_name, k)
        clusters = model.fit_predict(X)
        
        # 将聚类结果添加到数据中
        df_analysis = df.copy()
        df_analysis['cluster'] = clusters
        
        # 分析聚类特征
        cluster_features = []
        for i in range(k):
            cluster_data = df_analysis[df_analysis['cluster'] == i][selected_features]
            mean_values = cluster_data.mean()
            cluster_features.append(mean_values)
        
        # 创建热力图
        plt.figure(figsize=(15, 8))
        sns.heatmap(pd.DataFrame(cluster_features, columns=selected_features),
                   cmap='YlOrRd',
                   annot=True,
                   fmt='.2f',
                   cbar_kws={'label': '平均值'})
        plt.title(f'K={k} 各组的调味品使用特征分析')
        plt.xlabel('调味品')
        plt.ylabel('组别')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        # 保存热力图
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=300, bbox_inches='tight')
        
        # 添加图片到文档
        paragraph = doc.add_paragraph()
        # 设置段落格式
        paragraph_format = paragraph.paragraph_format
        paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph_format.space_before = Pt(12)
        paragraph_format.space_after = Pt(12)
        # 添加图片
        run = paragraph.add_run()
        run.add_picture(buf, width=Inches(7))
        plt.close()
        
        # 添加SKU占比表格标题
        title_paragraph = doc.add_paragraph('调味品使用占比（本类别中有使用的店数/本类别的总店数）')
        title_paragraph.alignment = 1
        title_paragraph.runs[0].bold = True
        
        # 创建SKU占比表格
        sku_table = doc.add_table(rows=k+1, cols=len(selected_features)+1)
        sku_table.style = 'Table Grid'
        
        # 设置表头
        header_cells = sku_table.rows[0].cells
        header_cells[0].text = '聚类'
        for idx, col in enumerate(selected_features):
            header_cells[idx+1].text = col
        
        # 设置表头样式
        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].alignment = 1
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
            tcPr.append(tcVAlign)
        
        # 填充表格内容
        for i in range(k):
            row_cells = sku_table.rows[i+1].cells
            
            # 聚类编号
            row_cells[0].text = f'聚类{i}'
            
            # 获取当前聚类的数据
            cluster_data = df_analysis[df_analysis['cluster'] == i]
            total_stores = len(cluster_data)
            
            # 计算每个调味品的使用比例
            for col_idx, col in enumerate(selected_features):
                stores_using_sku = (cluster_data[col] > 0).sum()
                if stores_using_sku > 0:
                    percentage = (stores_using_sku / total_stores) * 100
                    row_cells[col_idx+1].text = f'{int(round(percentage))}%'
                else:
                    row_cells[col_idx+1].text = '/'
        
        # 设置表格样式
        for row in sku_table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = 1
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
                tcPr.append(tcVAlign)
        
        doc.add_paragraph('')  # 添加空行
        
        if selected_cuisine == '全部' and selected_level == '全部':
            # 添加分布分析标题
            doc.add_heading(f'聚类 K={k} 的菜系和餐厅档次分布', level=2)
            
            # 为每个聚类进行菜系和餐厅档次分析
            for i in range(k):
                # 添加聚类标题
                doc.add_paragraph(f'聚类{i}')
                
                # 创建分布分析表格
                distribution_table = doc.add_table(rows=1, cols=3)
                distribution_table.style = 'Table Grid'
                
                # 设置表头
                header_cells = distribution_table.rows[0].cells
                header_cells[0].text = '类型'
                header_cells[1].text = '名称'
                header_cells[2].text = '占比'
                
                # 设置表头样式
                for cell in header_cells:
                    cell.paragraphs[0].runs[0].bold = True
                    cell.paragraphs[0].alignment = 1
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
                    tcPr.append(tcVAlign)
                
                # 获取当前聚类的数据
                cluster_data = df_analysis[df_analysis['cluster'] == i]
                
                # 菜系分布
                cuisine_dist = cluster_data[cuisine_col].value_counts()
                top_cuisines = cuisine_dist.head(3)
                
                # 餐厅档次分布
                level_dist = cluster_data[level_col].value_counts()
                
                # 添加菜系数据到表格
                for name, value in top_cuisines.items():
                    row = distribution_table.add_row()
                    cells = row.cells
                    cells[0].text = '菜系'
                    cells[1].text = str(name)
                    
                    # 计算百分比
                    total = len(cluster_data)
                    percentage = (value / total) * 100
                    cells[2].text = f'{int(round(percentage))}%'
                    
                    # 设置单元格样式
                    for cell in cells:
                        cell.paragraphs[0].alignment = 1
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
                        tcPr.append(tcVAlign)
                
                # 添加餐厅档次数据到表格
                for name, value in level_dist.items():
                    row = distribution_table.add_row()
                    cells = row.cells
                    cells[0].text = '餐厅档次'
                    cells[1].text = str(name)
                    
                    # 计算百分比
                    total = len(cluster_data)
                    percentage = (value / total) * 100
                    cells[2].text = f'{int(round(percentage))}%'
                    
                    # 设置单元格样式
                    for cell in cells:
                        cell.paragraphs[0].alignment = 1
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcVAlign = parse_xml(r'<w:vAlign xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
                        tcPr.append(tcVAlign)
                
                doc.add_paragraph('')  # 添加空行
        
        doc.add_paragraph('')  # 添加空行
    
    # 保存文档
    return save_report(doc, output_format)

def apply_document_style(doc):
    """统一文档样式"""
    # 设置默认字体
    style = doc.styles['Normal']
    style.font.name = '微软雅黑'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    # 设置标题样式
    for i in range(1, 4):
        style = doc.styles[f'Heading {i}']
        style.font.name = '微软雅黑'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        style.font.bold = True
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.font.size = Pt(16 - (i-1)*2)  # 标题逐级递减字号
    
    # 设置正文样式
    style = doc.styles['Normal']
    style.font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(10)

def save_report(doc, output_format='docx'):
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            docx_path = tmp_docx.name
        
        # 保存文档到临时文件
        doc.save(docx_path)
        
        if output_format == 'pdf':
            try:
                # 初始化 COM 环境
                pythoncom.CoInitialize()
                
                # 创建PDF临时文件路径
                pdf_path = docx_path.replace('.docx', '.pdf')
                # 转换为PDF
                convert(docx_path, pdf_path)
                
                # 读取PDF内容
                with open(pdf_path, 'rb') as pdf_file:
                    encoded_pdf = pdf_file.read()
                
                # 清理临时文件
                try:
                    os.remove(pdf_path)
                except:
                    pass
                try:
                    os.remove(docx_path)
                except:
                    pass
                
                # 取消初始化 COM 环境
                pythoncom.CoUninitialize()
                    
                return encoded_pdf, 'pdf'
                
            except Exception as e:
                # 确保在发生错误时也取消初始化 COM 环境
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
                    
                st.warning(f'PDF 转换失败，将使用 Word 格式: {str(e)}')
                # 如果PDF转换失败，返回Word文档
                with open(docx_path, 'rb') as docx_file:
                    encoded_docx = docx_file.read()
                try:
                    os.remove(docx_path)
                except:
                    pass
                return encoded_docx, 'docx'
        else:
            # 直接返回Word文档
            with open(docx_path, 'rb') as docx_file:
                encoded_docx = docx_file.read()
            try:
                os.remove(docx_path)
            except:
                pass
            return encoded_docx, 'docx'
    except Exception as e:
        st.error(f'保存报告时发生错误: {str(e)}')
        raise

def get_cuisine_types(df=None):
    """获取所有菜系类别，按照数据列D1到P1的顺序"""
    if df is not None:
        # 获取D到P列的列名（索引3到15）
        cuisine_columns = list(df.iloc[:, 3:16].columns)
        return ['全部'] + cuisine_columns
    return ['全部']  # 如果没有数据，只返回'全部'选项

def get_restaurant_levels():
    """获取餐厅档次类别"""
    return ['全部', '中高档', '大众']

def main():
    st.title('聚类分析报告生成器')
    
    # 上传文件
    uploaded_file = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        with st.spinner('正在处理数据，请稍候...'):
            all_data, sheet_names = load_data(uploaded_file)
        
        if all_data is not None and sheet_names:
            # 创建两列布局
            col_filters, col_preview = st.columns([1, 2])  # 1:2 的宽度比例
            
            # 左侧列：控件
            with col_filters:
                # Sheet选择
                selected_sheet = st.selectbox('选择工作表', sheet_names)
            
            if selected_sheet:
                df = all_data[selected_sheet]
                
                # 左侧列：继续添加控件
                with col_filters:
                    # 选择菜系和餐厅档次
                    cuisine_types = get_cuisine_types(df)  # 传入df以获取实际的列名
                    restaurant_levels = get_restaurant_levels()
                    
                    selected_cuisine = st.selectbox('选择菜系', cuisine_types, index=0)
                    selected_level = st.selectbox('选择餐厅档次', restaurant_levels, index=0)
                
                # 根据选择过滤数据
                filtered_df = df.copy()
                
                # 右侧列：数据预览
                with col_preview:
                    if selected_cuisine != '全部':
                        # 使用选定的列名进行过滤
                        filtered_df = filtered_df[filtered_df[selected_cuisine] > 0]
                    if selected_level != '全部':
                        filtered_df = filtered_df[filtered_df.iloc[:, 2] == selected_level]
                    
                    st.write("数据预览：")
                    st.write(filtered_df)
                    st.write(f"数据数量：{len(filtered_df)}")
                
                if len(filtered_df) == 0:
                    st.warning(f"没有找到符合条件的数据：{selected_cuisine} - {selected_level}")
                    st.stop()
                
                # 特征选择
                numeric_columns = list(df.select_dtypes(include=[np.number]).columns)
            
                # 当选择全部菜系和餐厅档次时，自动选择所有调味品列（D到P列）
                if selected_cuisine == '全部' and selected_level == '全部':
                    # 获取D到P列的列名
                    seasoning_columns = list(df.iloc[:, 3:16].columns)  # D列是index 3，P列是index 15
                    st.session_state.selected_features = seasoning_columns
                    selected_features = seasoning_columns
                
                    col3, col4 = st.columns(2)
                    with col3:
                        st.write("特征选择：")
                        st.write("已自动选择所有调味品特征 (D-P列)")
                    with col4:
                        # 选择聚类模型
                        model_name = st.selectbox(
                            '选择聚类模型',
                            ['KModes (适用于分类数据)', 'KMeans (适用于连续数据)', 'MiniBatchKMeans (适用于大规模数据)'],
                            index=0
                        )
                        model_name = model_name.split(' ')[0]  # 提取模型名称
                else:
                    col3, col4 = st.columns(2)
                    with col3:
                        st.write("特征选择：")
                        col3_1, col3_2 = st.columns(2)
                        with col3_1:
                            if st.button('全选特征'):
                                st.session_state.selected_features = numeric_columns
                        with col3_2:
                            if st.button('清除全部'):
                                st.session_state.selected_features = []
                    
                        if 'selected_features' not in st.session_state:
                            st.session_state.selected_features = numeric_columns[:5] if len(numeric_columns) > 5 else numeric_columns
                    
                        selected_features = st.multiselect(
                            '选择要分析的特征',
                            numeric_columns,
                            default=st.session_state.selected_features
                        )
                        st.session_state.selected_features = selected_features
                
                    with col4:
                        # 选择聚类模型
                        model_name = st.selectbox(
                            '选择聚类模型',
                            ['KModes (适用于分类数据)', 'KMeans (适用于连续数据)', 'MiniBatchKMeans (适用于大规模数据)'],
                            index=0
                        )
                        model_name = model_name.split(' ')[0]  # 提取模型名称
            
                if len(selected_features) > 0:
                    # 数据预处理
                    X = preprocess_data(filtered_df, selected_features)
                
                    # 计算评估指标
                    K, silhouette_scores, costs, optimal_k = find_optimal_k(X, 10)
                
                    # 显示评估图
                    st.write("聚类数量评估：")
                    k_value = st.slider('选择聚类数量', min_value=2, max_value=10, value=optimal_k)
                
                    # 创建评估图
                    fig = create_evaluation_plots(K, costs, silhouette_scores, k_value)
                    st.pyplot(fig)
                
                    # 显示评估说明
                    st.write("评估说明：")
                    st.write("1. 左图为肘部法则分析：当曲线开始平缓时的点为较优的聚类数量。")
                    st.write("2. 右图为轮廓系数分析：轮廓系数越大表示聚类效果越好。")
                    st.write(f"3. 根据综合评估，推荐的聚类数量为 {optimal_k}。")
                
                    col5, col6 = st.columns(2)
                    with col5:
                        st.write("报告生成范围：")
                        k_min = st.number_input('最小聚类数', min_value=2, max_value=9, value=4)
                        k_max = st.number_input('最大聚类数', min_value=3, max_value=10, value=min(k_min + 3, 10))
                
                    with col6:
                        # 输出格式选择
                        output_format = st.radio(
                            "选择输出格式",
                            ('docx', 'pdf'),
                            format_func=lambda x: 'Word文档' if x == 'docx' else 'PDF文档',
                            horizontal=True
                        )
                
                    if st.button('开始聚类分析'):
                        try:
                            with st.spinner('正在生成报告...'):
                                # 生成报告
                                k_range = range(k_min, k_max + 1)
                                encoded_report, actual_format = create_clustering_report(
                                    filtered_df, 
                                    f"{uploaded_file.name}_{selected_sheet}",
                                    1,  # 菜系列索引
                                    2,  # 餐厅档次列索引
                                    selected_features,
                                    k_range,
                                    model_name,
                                    output_format,
                                    selected_cuisine,  # 添加选择的菜系
                                    selected_level     # 添加选择的档次
                                )
                            
                                # 下载报告
                                extension = '.pdf' if actual_format == 'pdf' else '.docx'
                                filename = f'聚类分析报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}{extension}'
                            
                                st.download_button(
                                    label="下载报告",
                                    data=encoded_report,
                                    file_name=filename,
                                    mime=('application/pdf' if actual_format == 'pdf' 
                                         else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                                )
                            
                                st.success('报告生成成功！')
                        except Exception as e:
                            st.error(f'生成报告时发生错误: {str(e)}')
                            raise

if __name__ == '__main__':
    main()
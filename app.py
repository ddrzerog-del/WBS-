import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- 데이터 파싱 함수 ---
def parse_line(text):
    text = str(text).strip()
    match = re.match(r'^([\d\.]+)', text)
    if match:
        code = match.group(1).rstrip('.')
        level = code.count('.') + 1 # 1.1은 2레벨, 1.1.1은 3레벨
        return {'id_code': code, 'text': text, 'level': level}
    return None

# --- 트리 구조 구축 ---
def build_tree(data):
    nodes = {}
    root_nodes = []
    for item in data:
        code = item['id_code']
        # 노드 생성
        node = {'code': code, 'text': item['text'], 'level': item['level'], 'children': []}
        nodes[code] = node
        
        # 부모 찾기
        parts = code.split('.')
        if len(parts) > 1:
            parent_code = ".".join(parts[:-1])
            if parent_code in nodes:
                nodes[parent_code]['children'].append(node)
            else:
                # 부모가 아직 안 나타났거나 없는 경우 최상위로 (예외처리)
                if item['level'] == 1: root_nodes.append(node)
        else:
            root_nodes.append(node)
    return root_nodes

# --- 3레벨 이하 모든 자식을 리스트로 추출 (세로 나열용) ---
def get_all_descendants(node, desc_list):
    for child in node['children']:
        desc_list.append(child)
        get_all_descendants(child, desc_list)

# --- PPT 생성 함수 ---
def create_hybrid_wbs(root_nodes):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    margin_x = Inches(0.4)
    total_width = prs.slide_width - (margin_x * 2)
    
    # 레벨별 설정
    l1_height = Inches(0.6)
    l2_height = Inches(0.5)
    l3_plus_height = Inches(0.4)
    v_gap = Inches(0.15)
    
    # 1레벨 개수에 따라 가로 분할
    if not root_nodes: return prs
    l1_width = total_width / len(root_nodes)

    for i, l1 in enumerate(root_nodes):
        x_l1 = margin_x + (i * l1_width)
        y_l1 = Inches(0.5)
        
        # --- Level 1 그리기 ---
        shape1 = slide.shapes.add_shape(1, x_l1, y_l1, l1_width - Inches(0.1), l1_height)
        shape1.fill.solid()
        shape1.fill.fore_color.rgb = RGBColor(31, 73, 125) # 진한 파랑
        tf1 = shape1.text_frame
        tf1.text = l1['text']
        tf1.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf1.paragraphs[0].font.size = Pt(11)
        tf1.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        tf1.paragraphs[0].font.bold = True

        if l1['children']:
            # 2레벨 가로 너비 (1레벨 너비 내에서 분할)
            l2_width = (l1_width - Inches(0.1)) / len(l1['children'])
            
            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * l2_width)
                y_l2 = y_l1 + l1_height + v_gap
                
                # --- Level 2 그리기 ---
                shape2 = slide.shapes.add_shape(1, x_l2, y_l2, l2_width - Inches(0.05), l2_height)
                shape2.fill.solid()
                shape2.fill.fore_color.rgb = RGBColor(54, 95, 145) # 중간 파랑
                tf2 = shape2.text_frame
                tf2.text = l2['text']
                tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
                tf2.paragraphs[0].font.size = Pt(10)
                tf2.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                # --- Level 3 이하 세로 나열 ---
                descendants = []
                get_all_descendants(l2, descendants)
                
                current_y_l3 = y_l2 + l2_height + v_gap
                for k, desc in enumerate(descendants):
                    # 텍스트가 너무 많으면 슬라이드를 넘어갈 수 있으므로 높이 조절
                    shape3 = slide.shapes.add_shape(1, x_l2, current_y_l3, l2_width - Inches(0.05), l3_plus_height)
                    shape3.fill.solid()
                    
                    # 레벨이 깊어질수록 연한 색상
                    color_val = min(150 + (desc['level'] * 20), 240)
                    shape3.fill.fore_color.rgb = RGBColor(color_val, color_val, color_val + 10)
                    shape3.line.color.rgb = RGBColor(200, 200, 200)
                    
                    tf3 = shape3.text_frame
                    tf3.text = desc['text']
                    tf3.paragraphs[0].alignment = PP_ALIGN.LEFT
                    p3 = tf3.paragraphs[0]
                    p3.font.size = Pt(8)
                    p3.font.color.rgb = RGBColor(0, 0, 0)
                    
                    # 다음 박스 위치 (누적)
                    current_y_l3 += l3_plus_height + Inches(0.05)

    return prs

# --- Streamlit UI ---
st.set_page_config(page_title="WBS Hybrid Aligner", layout="wide")
st.title("📊 하이브리드형 WBS 자동 생성기")
st.subheader("1-2단계는 가로로, 3단계 이하는 세로로 자동 정렬합니다.")

uploaded_file = st.file_uploader("엑셀(.xlsx) 또는 파워포인트(.pptx) 파일 업로드", type=["xlsx", "pptx"])

if uploaded_file:
    raw_data = []
    if uploaded_file.name.endswith("xlsx"):
        df = pd.read_excel(uploaded_file)
        for val in df.iloc[:, 0]:
            p = parse_line(val)
            if p: raw_data.append(p)
    else:
        input_prs = Presentation(uploaded_file)
        for s in input_prs.slides:
            for shp in s.shapes:
                if hasattr(shp, "text"):
                    p = parse_line(shp.text)
                    if p: raw_data.append(p)

    if raw_data:
        # ID 순서로 정렬
        raw_data.sort(key=lambda x: [int(i) for i in x['id_code'].split('.')])
        tree = build_tree(raw_data)
        
        st.success(f"데이터 로드 완료: {len(raw_data)}개 항목 인식")
        
        if st.button("🚀 하이브리드 WBS 생성"):
            final_ppt = create_hybrid_wbs(tree)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 완성된 PPT 다운로드", ppt_io, "Hybrid_WBS.pptx")

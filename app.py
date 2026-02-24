import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- 데이터 파싱 및 트리 구축 ---
def parse_line(text):
    text = str(text).strip()
    match = re.match(r'^([\d\.]+)', text)
    if match:
        code = match.group(1).rstrip('.')
        level = code.count('.') + 1
        return {'id_code': code, 'text': text, 'level': level}
    return None

def build_tree(data):
    nodes = {}
    root_nodes = []
    for item in data:
        code = item['id_code']
        node = {'code': code, 'text': item['text'], 'level': item['level'], 'children': []}
        nodes[code] = node
        parts = code.split('.')
        if len(parts) > 1:
            parent_code = ".".join(parts[:-1])
            if parent_code in nodes:
                nodes[parent_code]['children'].append(node)
            else:
                if item['level'] == 1: root_nodes.append(node)
        else:
            root_nodes.append(node)
    return root_nodes

def get_all_descendants(node, desc_list):
    for child in node['children']:
        desc_list.append(child)
        get_all_descendants(child, desc_list)

# --- PPT 생성 (고도화 레이아웃) ---
def create_advanced_wbs(root_nodes):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 설정 상수 (Inches 단위)
    MARGIN_X = Inches(0.5)
    TOTAL_WIDTH = prs.slide_width - (2 * MARGIN_X)
    
    BASE_V_GAP = 0.3  # 기본 수직 간격
    WIDTH_STEP = 0.15 # 레벨당 줄어들 너비 (Inches)
    GROUP_GAP = Inches(0.2) # 그룹 간 물리적 이격

    if not root_nodes: return prs

    # 1레벨 너비 계산 (그룹 간 여백 포함)
    l1_count = len(root_nodes)
    l1_width_with_gap = TOTAL_WIDTH / l1_count
    l1_width = l1_width_with_gap - GROUP_GAP

    for i, l1 in enumerate(root_nodes):
        # 1레벨 시작 X
        x_l1_start = MARGIN_X + (i * l1_width_with_gap)
        y_l1 = Inches(0.6)
        l1_h = Inches(0.6)

        # 1레벨 상자
        shp1 = slide.shapes.add_shape(1, x_l1_start, y_l1, l1_width, l1_h)
        shp1.fill.solid()
        shp1.fill.fore_color.rgb = RGBColor(31, 73, 125)
        shp1.text = l1['text']
        shp1.text_frame.paragraphs[0].font.size = Pt(11)
        shp1.text_frame.paragraphs[0].font.bold = True

        if l1['children']:
            # 2레벨 너비 (1레벨 박스 안에서 분할)
            l2_count = len(l1['children'])
            l2_width_full = l1_width / l2_count
            l2_width = l2_width_full - Inches(0.05) # 2레벨간 미세 간격
            
            # 수직 간격 (1-2레벨 간격은 10 비율)
            v_gap_l1_l2 = Inches(BASE_V_GAP)

            for j, l2 in enumerate(l1['children']):
                x_l2_start = x_l1_start + (j * l2_width_full)
                y_l2 = y_l1 + l1_h + v_gap_l1_l2
                l2_h = Inches(0.5)

                # 2레벨 상자
                shp2 = slide.shapes.add_shape(1, x_l2_start, y_l2, l2_width, l2_h)
                shp2.fill.solid()
                shp2.fill.fore_color.rgb = RGBColor(54, 95, 145)
                shp2.text = l2['text']
                shp2.text_frame.paragraphs[0].font.size = Pt(10)

                # 3레벨 이하 (우측 정렬 및 계단식 너비/간격)
                descendants = []
                get_all_descendants(l2, descendants)
                
                current_y = y_l2 + l2_h
                
                # 레벨별 상대적 좌표 계산용
                for k, desc in enumerate(descendants):
                    # 1. 수직 간격 차등화 (깊어질수록 좁아짐: 10, 9, 8...)
                    # 3레벨 이상부터는 조금씩 더 좁게 배치
                    gap_factor = max(0.5, 1.0 - (desc['level'] - 2) * 0.1)
                    current_v_gap = Inches(BASE_V_GAP * 0.7 * gap_factor)
                    current_y += current_v_gap

                    # 2. 박스 너비 계단식 차이 (L2 대비 4.9, 4.8...)
                    reduction = Inches(WIDTH_STEP * (desc['level'] - 2))
                    desc_width = l2_width - reduction
                    if desc_width < Inches(1.0): desc_width = Inches(1.0) # 최소 너비 보장

                    # 3. 우측 끝 정렬 (Parent Right - My Width)
                    parent_right_x = x_l2_start + l2_width
                    desc_left_x = parent_right_x - desc_width

                    # 4. 상자 그리기
                    desc_h = Inches(0.4)
                    shp_d = slide.shapes.add_shape(1, desc_left_x, current_y, desc_width, desc_h)
                    
                    # 디자인: 레벨이 깊을수록 연해짐
                    c_val = min(180 + (desc['level'] * 15), 245)
                    shp_d.fill.solid()
                    shp_d.fill.fore_color.rgb = RGBColor(c_val, c_val, c_val + 5)
                    shp_d.line.color.rgb = RGBColor(200, 200, 200)
                    
                    shp_d.text = desc['text']
                    tf = shp_d.text_frame
                    tf.paragraphs[0].font.size = Pt(8)
                    tf.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
                    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
                    
                    # Y축 업데이트
                    current_y += desc_h

    return prs

# --- Streamlit UI ---
st.set_page_config(page_title="Advanced WBS Aligner", layout="wide")
st.title("🚀 고도화된 WBS 자동 정렬기")
st.markdown("""
- **그룹화**: 레벨 1/2 간 그룹 여백 적용
- **계단식 디자인**: 하위 레벨로 갈수록 박스 크기와 간격이 미세하게 축소
- **우측 정렬**: 하위 항목들이 부모 항목의 우측 끝 라인에 맞춰 정렬
""")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, pptx)", type=["xlsx", "pptx"])

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
        raw_data.sort(key=lambda x: [int(i) for i in x['id_code'].split('.')])
        tree = build_tree(raw_data)
        
        if st.button("🎨 고도화 PPT 생성"):
            final_ppt = create_advanced_wbs(tree)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 정렬된 PPT 다운로드", ppt_io, "Advanced_WBS.pptx")

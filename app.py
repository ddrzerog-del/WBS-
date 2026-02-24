import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- 데이터 파싱 및 트리 구축 (이전과 동일) ---
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

# --- 메인 PPT 생성 함수 ---
def create_final_wbs(root_nodes, config):
    prs = Presentation()
    # 슬라이드 크기 설정 (사용자 입력에 따라 유동적일 수 있으나 기본 16:9 권장)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 설정값 (cm -> pptx 내부 단위 변환)
    wbs_w = Cm(config['wbs_w_cm'])
    wbs_h = Cm(config['wbs_h_cm'])
    l1_gap = Cm(config['l1_gap_cm'])
    l2_gap = Cm(config['l2_gap_cm'])
    v_gap_base = Cm(config['v_gap_cm'])
    
    # 시작 좌표 (가운데 정렬을 위해 계산)
    start_x = (prs.slide_width - wbs_w) / 2
    start_y = (prs.slide_height - wbs_h) / 2

    if not root_nodes: return prs

    # 1레벨 박스 너비 계산
    l1_count = len(root_nodes)
    # 전체너비 = (l1_width * l1_count) + (l1_gap * (l1_count - 1))
    l1_width = (wbs_w - (l1_gap * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        curr_l1_x = start_x + (i * (l1_width + l1_gap))
        l1_h = Cm(1.2) # 1레벨 높이는 고정 권장
        
        # 1레벨 상자
        shp1 = slide.shapes.add_shape(1, curr_l1_x, start_y, l1_width, l1_h)
        shp1.fill.solid()
        shp1.fill.fore_color.rgb = RGBColor(31, 73, 125)
        shp1.text = l1['text']
        shp1.text_frame.paragraphs[0].font.size = Pt(12)
        shp1.text_frame.paragraphs[0].font.bold = True

        if l1['children']:
            l2_count = len(l1['children'])
            # 2레벨 너비 (1레벨 박스 영역 내에서 계산)
            l2_width = (l1_width - (l2_gap * (l2_count - 1))) / l2_count
            
            for j, l2 in enumerate(l1['children']):
                curr_l2_x = curr_l1_x + (j * (l2_width + l2_gap))
                y_l2 = start_y + l1_h + v_gap_base
                l2_h = Cm(1.0)

                # 2레벨 상자
                shp2 = slide.shapes.add_shape(1, curr_l2_x, y_l2, l2_width, l2_h)
                shp2.fill.solid()
                shp2.fill.fore_color.rgb = RGBColor(54, 95, 145)
                shp2.text = l2['text']
                shp2.text_frame.paragraphs[0].font.size = Pt(10)

                # 3레벨 이하 상세항목
                descendants = []
                get_all_descendants(l2, descendants)
                
                current_y = y_l2 + l2_h
                for k, desc in enumerate(descendants):
                    # 레벨에 따른 수직 간격 및 너비 계단식 축소
                    step_v_gap = v_gap_base * 0.6 * (0.9 ** (desc['level'] - 3))
                    current_y += step_v_gap
                    
                    # 너비 축소 (Cm(0.2)씩 계단식 축소)
                    reduction = Cm(0.3 * (desc['level'] - 2))
                    desc_w = l2_width - reduction
                    if desc_w < Cm(2.0): desc_w = Cm(2.0) # 최소 크기 방어선

                    # 우측 정렬
                    parent_right = curr_l2_x + l2_width
                    desc_x = parent_right - desc_w
                    
                    desc_h = Cm(0.8)
                    shp_d = slide.shapes.add_shape(1, desc_x, current_y, desc_w, desc_h)
                    
                    # 색상 및 텍스트 설정
                    c_val = min(190 + (desc['level'] * 15), 245)
                    shp_d.fill.solid()
                    shp_d.fill.fore_color.rgb = RGBColor(c_val, c_val, c_val + 10)
                    shp_d.line.color.rgb = RGBColor(200, 200, 200)
                    shp_d.text = desc['text']
                    
                    tf = shp_d.text_frame
                    tf.paragraphs[0].font.size = Pt(8)
                    tf.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
                    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
                    
                    current_y += desc_h

    return prs

# --- Streamlit UI ---
st.set_page_config(page_title="WBS Custom Aligner", layout="wide")

# 사이드바 설정창
st.sidebar.header("🎨 디자인 옵션")

st.sidebar.subheader("1. 전체 영역 크기 (cm)")
wbs_w_cm = st.sidebar.number_input("WBS 전체 너비", value=30.0, step=1.0)
wbs_h_cm = st.sidebar.number_input("WBS 전체 높이", value=15.0, step=1.0)

st.sidebar.subheader("2. 간격 조절 (cm)")
l1_gap_cm = st.sidebar.slider("대그룹(L1) 좌우 간격", 0.0, 5.0, 1.5)
l2_gap_cm = st.sidebar.slider("소그룹(L2) 좌우 간격", 0.0, 3.0, 0.5)
v_gap_cm = st.sidebar.slider("상하(Vertical) 기본 간격", 0.1, 2.0, 0.5)

config = {
    'wbs_w_cm': wbs_w_cm, 'wbs_h_cm': wbs_h_cm,
    'l1_gap_cm': l1_gap_cm, 'l2_gap_cm': l2_gap_cm, 'v_gap_cm': v_gap_cm
}

st.title("📊 커스텀 WBS 자동 정렬 프로그램")
st.write("사이드바에서 간격과 크기를 조절한 후 PPT를 생성하세요.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, pptx)", type=["xlsx", "pptx"])

if uploaded_file:
    raw_data = []
    # 데이터 읽기 (생략 - 이전과 동일)
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
        
        st.info(f"선택한 영역: {wbs_w_cm}cm x {wbs_h_cm}cm")
        
        if st.button("🚀 설정값으로 PPT 생성"):
            final_ppt = create_final_wbs(tree, config)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 완성된 PPT 다운로드", ppt_io, "Custom_WBS.pptx")

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- WBS 트리 노드 클래스 ---
class WBSNode:
    def __init__(self, id_code, text, level):
        self.id_code = id_code
        self.text = text
        self.level = level
        self.children = []
        self.width_factor = 1  # 이 노드가 차지할 가로 비중
        self.x_pos = 0        # 계산된 X 좌표
        self.final_width = 0  # 계산된 실제 너비

# --- 트리 생성 함수 ---
def build_tree(data):
    nodes = {}
    root_nodes = []
    
    # 1. 노드 객체 생성
    for item in data:
        code = item['id_code']
        node = WBSNode(code, item['text'], item['level'])
        nodes[code] = node
        
        # 부모 찾기 (예: 1.1.1의 부모는 1.1)
        parent_code = ".".join(code.split(".")[:-1])
        if parent_code in nodes:
            nodes[parent_code].children.append(node)
        else:
            root_nodes.append(node)
    
    # 2. 너비 계수 계산 (Bottom-up)
    def calc_width_factor(node):
        if not node.children:
            node.width_factor = 1
            return 1
        factor = sum(calc_width_factor(child) for child in node.children)
        node.width_factor = max(factor, 1)
        return node.width_factor

    for root in root_nodes:
        calc_width_factor(root)
        
    return root_nodes

# --- 텍스트 파싱 로직 ---
def parse_line(text):
    text = str(text).strip()
    match = re.match(r'^([\d\.]+)', text)
    if match:
        code = match.group(1).rstrip('.')
        level = code.count('.')
        return {'id_code': code, 'text': text, 'level': level}
    return None

# --- PPT 생성 로직 ---
def create_wbs_ppt(root_nodes):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    margin_x = Inches(0.5)
    total_width = prs.slide_width - (margin_x * 2)
    box_height = Inches(0.6)
    v_gap = Inches(0.3)
    
    total_factors = sum(root.width_factor for root in root_nodes)
    unit_width = total_width / total_factors

    # 레벨별 색상 (진한색 -> 연한색)
    colors = [RGBColor(31, 73, 125), RGBColor(54, 95, 145), RGBColor(79, 129, 189), RGBColor(149, 179, 215), RGBColor(198, 217, 241)]

    # 재귀적으로 그리기
    def draw_node(node, current_x, current_y):
        node_width = node.width_factor * unit_width
        
        # 도형 그리기
        shape = slide.shapes.add_shape(
            1, current_x, current_y, node_width - Inches(0.05), box_height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = colors[min(node.level, len(colors)-1)]
        shape.line.color.rgb = RGBColor(255, 255, 255)
        
        tf = shape.text_frame
        tf.text = node.text
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        p = tf.paragraphs[0]
        p.font.size = Pt(9)
        p.font.color.rgb = RGBColor(255, 255, 255)
        
        # 자식 노드 배치
        child_x = current_x
        for child in node.children:
            draw_node(child, child_x, current_y + box_height + v_gap)
            child_x += (child.width_factor * unit_width)

    start_x = margin_x
    for root in root_nodes:
        draw_node(root, start_x, Inches(0.5))
        start_x += (root.width_factor * unit_width)

    return prs

# --- Streamlit UI ---
st.set_page_config(page_title="WBS Pro Aligner", layout="wide")
st.title("📂 하이브리드 WBS 자동 정렬기")
st.write("계층 구조를 분석하여 부모 항목 아래에 자식 항목들을 완벽하게 그룹화하여 정렬합니다.")

uploaded_file = st.file_uploader("엑셀 또는 PPT 파일 업로드", type=["xlsx", "pptx"])

if uploaded_file:
    raw_data = []
    if uploaded_file.name.endswith("xlsx"):
        df = pd.read_excel(uploaded_file)
        for val in df.iloc[:, 0]:
            p = parse_line(val)
            if p: raw_data.append(p)
    else:
        # PPT 처리 로직 생략(위와 동일)
        prs_in = Presentation(uploaded_file)
        for s in prs_in.slides:
            for shp in s.shapes:
                if hasattr(shp, "text"):
                    p = parse_line(shp.text)
                    if p: raw_data.append(p)

    if raw_data:
        # ID 순서로 정렬 (1, 1.1, 1.1.1 순)
        raw_data.sort(key=lambda x: [int(i) for i in x['id_code'].split('.')])
        
        root_nodes = build_tree(raw_data)
        st.success(f"데이터 트리 구조 생성 완료 ({len(raw_data)}개 항목)")

        if st.button("🚀 그룹화 정렬 PPT 생성"):
            out_prs = create_wbs_ppt(root_nodes)
            ppt_io = io.BytesIO()
            out_prs.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 완성된 PPT 다운로드", ppt_io, "Smart_WBS.pptx")

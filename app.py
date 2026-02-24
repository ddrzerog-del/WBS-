import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# --- 1. 데이터 파싱 및 트리 구조화 ---
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

# --- 2. 좌표 계산 로직 (미리보기 & PPT 공용) ---
# 모든 노드의 x, y, width, height를 cm 단위로 미리 계산합니다.
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w = config['wbs_w']
    wbs_h = config['wbs_h']
    l1_gap = config['l1_gap']
    l2_gap = config['l2_gap']
    v_gap = config['v_gap']
    
    # 시작 원점 (중앙 정렬용)
    # 실제 PPT 슬라이드 크기(16:9)는 약 33.8cm x 19.05cm
    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2

    l1_count = len(root_nodes)
    if l1_count == 0: return []
    
    l1_width = (wbs_w - (l1_gap * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        x_l1 = start_x + (i * (l1_width + l1_gap))
        y_l1 = start_y
        h_l1 = 1.2
        layout_data.append({'node': l1, 'x': x_l1, 'y': y_l1, 'w': l1_width, 'h': h_l1, 'level': 1})

        if l1['children']:
            l2_count = len(l1['children'])
            l2_width = (l1_width - (l2_gap * (l2_count - 1))) / l2_count
            
            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap))
                y_l2 = y_l1 + h_l1 + v_gap
                h_l2 = 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})

                descendants = []
                get_all_descendants(l2, descendants)
                curr_y = y_l2 + h_l2
                
                for k, desc in enumerate(descendants):
                    # 간격 및 너비 축소 적용
                    step_v = v_gap * 0.6 * (0.9 ** (desc['level'] - 3))
                    curr_y += step_v
                    
                    reduction = 0.4 * (desc['level'] - 2)
                    d_w = max(l2_width - reduction, 2.0)
                    d_x = (x_l2 + l2_width) - d_w # 우측 정렬
                    d_h = 0.8
                    
                    layout_data.append({'node': desc, 'x': d_x, 'y': curr_y, 'w': d_w, 'h': d_h, 'level': desc['level']})
                    curr_y += d_h
                    
    return layout_data

# --- 3. 미리보기 (Matplotlib) ---
def draw_preview(layout_data):
    fig, ax = plt.subplots(figsize=(12, 6.75)) # 16:9 비율
    ax.set_xlim(0, 33.8)
    ax.set_ylim(0, 19.05)
    ax.invert_yaxis() # PPT처럼 상단이 0
    
    # 슬라이드 테두리
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, linewidth=1, edgecolor='black', facecolor='#f0f0f0', alpha=0.3))

    for item in layout_data:
        lvl = item['level']
        # 레벨별 색상 설정
        color = '#1f497d' if lvl == 1 else '#365f91' if lvl == 2 else '#d9d9d9'
        rect = patches.Rectangle((item['x'], item['y']), item['w'], item['h'], 
                                 linewidth=1, edgecolor='white', facecolor=color)
        ax.add_patch(rect)
        
        # 텍스트 요약 (너무 길면 자름)
        display_text = item['node']['text'][:10] + ".." if len(item['node']['text']) > 10 else item['node']['text']
        txt_color = 'white' if lvl <= 2 else 'black'
        ax.text(item['x'] + item['w']/2, item['y'] + item['h']/2, display_text, 
                color=txt_color, fontsize=7, ha='center', va='center')

    ax.set_axis_off()
    st.pyplot(fig)

# --- 4. PPT 생성 ---
def generate_ppt(layout_data):
    prs = Presentation()
    prs.slide_width = Cm(33.8)
    prs.slide_height = Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for item in layout_data:
        shp = slide.shapes.add_shape(1, Cm(item['x']), Cm(item['y']), Cm(item['w']), Cm(item['h']))
        lvl = item['level']
        
        # 디자인 적용
        shp.fill.solid()
        if lvl == 1:
            shp.fill.fore_color.rgb = RGBColor(31, 73, 125)
            font_size, font_bold, font_color = Pt(12), True, RGBColor(255, 255, 255)
        elif lvl == 2:
            shp.fill.fore_color.rgb = RGBColor(54, 95, 145)
            font_size, font_bold, font_color = Pt(10), False, RGBColor(255, 255, 255)
        else:
            c = min(200 + (lvl * 10), 245)
            shp.fill.fore_color.rgb = RGBColor(c, c, c+5)
            shp.line.color.rgb = RGBColor(200, 200, 200)
            font_size, font_bold, font_color = Pt(8), False, RGBColor(0, 0, 0)
            
        tf = shp.text_frame
        tf.text = item['node']['text']
        p = tf.paragraphs[0]
        p.font.size = font_size
        p.font.bold = font_bold
        p.font.color.rgb = font_color
        p.alignment = PP_ALIGN.CENTER if lvl <= 2 else PP_ALIGN.LEFT
        
    return prs

# --- 5. Streamlit UI ---
st.set_page_config(page_title="WBS Designer Pro", layout="wide")

st.sidebar.title("🎨 WBS 상세 설정")

# 사이드바: 수치 입력창 (number_input 사용)
with st.sidebar.expander("📏 전체 크기 설정 (cm)", expanded=True):
    wbs_w = st.number_input("WBS 전체 가로 너비", 10.0, 32.0, 30.0, 0.5)
    wbs_h = st.number_input("WBS 전체 세로 높이", 5.0, 18.0, 15.0, 0.5)

with st.sidebar.expander("↔️ 간격 설정 (cm)", expanded=True):
    l1_gap = st.number_input("대그룹(L1) 간격", 0.0, 10.0, 1.5, 0.1)
    l2_gap = st.number_input("소그룹(L2) 간격", 0.0, 10.0, 0.5, 0.1)
    v_gap = st.number_input("상하(Vertical) 기본 간격", 0.0, 5.0, 0.5, 0.05)

config = {'wbs_w': wbs_w, 'wbs_h': wbs_h, 'l1_gap': l1_gap, 'l2_gap': l2_gap, 'v_gap': v_gap}

st.title("📊 WBS 프로 디자이너")
st.write("엑셀/PPT를 업로드하고 왼쪽 설정창에서 수치를 변경하면 실시간으로 미리보기가 업데이트됩니다.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, pptx)", type=["xlsx", "pptx"])

if uploaded_file:
    # 데이터 파싱
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
        
        # 레이아웃 계산
        layout_data = calculate_layout(tree, config)
        
        # 미리보기 영역
        st.subheader("🖼️ 슬라이드 미리보기")
        draw_preview(layout_data)
        
        # 하단 다운로드 버튼
        st.divider()
        col1, col2 = st.columns([4, 1])
        with col2:
            if st.button("🚀 최종 PPT 생성 및 다운로드", use_container_width=True):
                final_ppt = generate_ppt(layout_data)
                ppt_io = io.BytesIO()
                final_ppt.save(ppt_io)
                ppt_io.seek(0)
                st.download_button("🎁 PPT 파일 받기", ppt_io, "Smart_WBS_Final.pptx")

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

# --- [캐싱 적용] 데이터 파싱 함수 ---
@st.cache_data
def process_uploaded_file(file_content, file_name):
    raw_data = []
    
    # 텍스트 파싱 내부 함수
    def parse_text(text):
        text = str(text).strip()
        match = re.match(r'^([\d\.]+)', text)
        if match:
            code = match.group(1).rstrip('.')
            level = code.count('.') + 1
            return {'id_code': code, 'text': text, 'level': level}
        return None

    if file_name.endswith("xlsx"):
        df = pd.read_excel(file_content)
        for val in df.iloc[:, 0]:
            p = parse_text(val)
            if p: raw_data.append(p)
    else:
        input_prs = Presentation(file_content)
        for s in input_prs.slides:
            for shp in s.shapes:
                if hasattr(shp, "text"):
                    p = parse_text(shp.text)
                    if p: raw_data.append(p)
                    
    # 정렬까지 마친 상태로 반환
    raw_data.sort(key=lambda x: [int(i) for i in x['id_code'].split('.')])
    return raw_data

# --- [캐싱 적용] 트리 구축 함수 ---
@st.cache_data
def build_wbs_tree(raw_data):
    nodes = {}
    root_nodes = []
    for item in raw_data:
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

# --- 좌표 계산 및 기타 함수는 동일하게 유지 ---
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w, wbs_h = config['wbs_w'], config['wbs_h']
    l1_gap_x, l2_gap_x = config['l1_gap_x'], config['l2_gap_x']
    v_gap_a = config['v_gap_a']
    extra_gaps = {3: config['extra_l3'], 4: config['extra_l4'], 5: config['extra_l5']}

    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2
    l1_count = len(root_nodes)
    if l1_count == 0: return []
    l1_width = (wbs_w - (l1_gap_x * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        x_l1, y_l1, h_l1 = start_x + (i * (l1_width + l1_gap_x)), start_y, 1.0
        layout_data.append({'node': l1, 'x': x_l1, 'y': y_l1, 'w': l1_width, 'h': h_l1, 'level': 1})
        if l1['children']:
            l2_width = (l1_width - (l2_gap_x * (len(l1['children']) - 1))) / len(l1['children'])
            y_l2_start = y_l1 + h_l1 + v_gap_a
            for j, l2 in enumerate(l1['children']):
                x_l2, y_l2, h_l2 = x_l1 + (j * (l2_width + l2_gap_x)), y_l2_start, 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})
                def stack_recursive(parent_node, px, py, pw, ph):
                    nonlocal layout_data
                    last_y = py + ph
                    for idx, child in enumerate(parent_node['children']):
                        gap = v_gap_a + (extra_gaps.get(child['level'], extra_gaps[5] if child['level']>5 else 0) if idx > 0 else 0)
                        target_y = last_y + gap
                        reduction = 0.3 * (child['level'] - 2)
                        c_w, c_h = max(pw - reduction, 2.0), 0.8
                        c_x = (px + pw) - c_w
                        layout_data.append({'node': child, 'x': c_x, 'y': target_y, 'w': c_w, 'h': c_h, 'level': child['level']})
                        last_y = stack_recursive(child, c_x, target_y, c_w, c_h)
                    return last_y
                stack_recursive(l2, x_l2, y_l2, l2_width, h_l2)
    return layout_data

def draw_preview(layout_data):
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.set_xlim(0, 33.8); ax.set_ylim(0, 19.05); ax.invert_yaxis()
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, facecolor='#f8f9fa', edgecolor='#dee2e6'))
    for item in layout_data:
        lvl = item['level']
        color = '#1f497d' if lvl == 1 else '#365f91' if lvl == 2 else '#f1f3f5'
        rect = patches.Rectangle((item['x'], item['y']), item['w'], item['h'], linewidth=0.5, edgecolor='#adb5bd', facecolor=color)
        ax.add_patch(rect)
        txt_color = 'white' if lvl <= 2 else 'black'
        ax.text(item['x'] + item['w']/2, item['y'] + item['h']/2, item['node']['text'][:12], color=txt_color, fontsize=6, ha='center', va='center')
    ax.set_axis_off()
    st.pyplot(fig)

def generate_ppt(layout_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.8), Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for item in layout_data:
        shp = slide.shapes.add_shape(1, Cm(item['x']), Cm(item['y']), Cm(item['w']), Cm(item['h']))
        lvl = item['level']
        shp.fill.solid()
        if lvl == 1: shp.fill.fore_color.rgb = RGBColor(31, 73, 125); f_size, f_bold, f_color, align = Pt(12), True, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        elif lvl == 2: shp.fill.fore_color.rgb = RGBColor(54, 95, 145); f_size, f_bold, f_color, align = Pt(10), False, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        else:
            c = min(220 + (lvl * 5), 250); shp.fill.fore_color.rgb = RGBColor(c, c, c+5); shp.line.color.rgb = RGBColor(200, 200, 200)
            f_size, f_bold, f_color, align = Pt(8.5), False, RGBColor(0, 0, 0), PP_ALIGN.LEFT
        tf = shp.text_frame; tf.text = item['node']['text']
        p = tf.paragraphs[0]; p.font.size, p.font.bold, p.font.color.rgb, p.alignment = f_size, f_bold, f_color, align
    return prs

# --- [UI] 사이드바 및 메인 ---
st.set_page_config(page_title="WBS Fast Designer", layout="wide")

st.sidebar.title("🎨 WBS 프로 디자인 설정")

# 사이드바 입력창들
with st.sidebar.expander("📏 캔버스 크기 (cm)", expanded=True):
    wbs_w = st.number_input("WBS 전체 너비", 10.0, 32.0, 31.0, key="w_in")
    wbs_h = st.number_input("WBS 전체 높이", 5.0, 18.0, 16.0, key="h_in")

with st.sidebar.expander("↕️ 상하 간격 정밀 설정 (cm)", expanded=True):
    v_gap_a = st.number_input("기준 수직 간격 (A)", 0.0, 5.0, 0.4, 0.05, key="a_in")
    extra_l3 = st.number_input("L3 그룹 간 추가 여백", 0.0, 5.0, 0.3, 0.05, key="l3_in")
    extra_l4 = st.number_input("L4 그룹 간 추가 여백", 0.0, 5.0, 0.2, 0.05, key="l4_in")
    extra_l5 = st.number_input("L5+ 그룹 간 추가 여백", 0.0, 5.0, 0.1, 0.05, key="l5_in")

with st.sidebar.expander("↔️ 좌우 간격 설정 (cm)", expanded=True):
    l1_gap_x = st.number_input("L1 좌우 간격", 0.0, 10.0, 1.2, key="l1_in")
    l2_gap_x = st.number_input("L2 좌우 간격", 0.0, 5.0, 0.4, key="l2_in")

config = {'wbs_w': wbs_w, 'wbs_h': wbs_h, 'l1_gap_x': l1_gap_x, 'l2_gap_x': l2_gap_x, 'v_gap_a': v_gap_a, 'extra_l3': extra_l3, 'extra_l4': extra_l4, 'extra_l5': extra_l5}

st.title("📊 WBS 프로 디자이너 (최적화 버전)")

uploaded_file = st.file_uploader("파일 업로드", type=["xlsx", "pptx"])

if uploaded_file:
    # 캐싱된 함수 호출: 파일 이름과 내용을 기반으로 결과 저장
    # 파일이 바뀌지 않으면 아래 작업은 0.01초 만에 끝납니다.
    raw_data = process_uploaded_file(uploaded_file.getvalue(), uploaded_file.name)
    
    if raw_data:
        tree = build_wbs_tree(raw_data)
        
        # 레이아웃 계산 및 미리보기는 캐싱하지 않음 (실시간 반영 필요)
        layout_data = calculate_layout(tree, config)
        
        st.subheader("🖼️ 디자인 미리보기")
        draw_preview(layout_data)
        
        if st.button("🚀 PPT 생성 및 다운로드", use_container_width=True):
            final_ppt = generate_ppt(layout_data)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 PPT 파일 다운로드", ppt_io, "Smart_WBS_Final.pptx")

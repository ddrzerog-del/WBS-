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

# --- 2. 좌표 계산 로직 (통합 간격 + 레벨별 그룹 추가 여백) ---
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w = config['wbs_w']
    wbs_h = config['wbs_h']
    l1_gap_x = config['l1_gap_x']
    l2_gap_x = config['l2_gap_x']
    
    # 간격 설정
    v_gap_a = config['v_gap_a']
    extra_gaps = {
        3: config['extra_l3'],
        4: config['extra_l4'],
        5: config['extra_l5']
    }

    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2

    l1_count = len(root_nodes)
    if l1_count == 0: return []
    l1_width = (wbs_w - (l1_gap_x * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        x_l1 = start_x + (i * (l1_width + l1_gap_x))
        y_l1 = start_y
        h_l1 = 1.0
        layout_data.append({'node': l1, 'x': x_l1, 'y': y_l1, 'w': l1_width, 'h': h_l1, 'level': 1})

        if l1['children']:
            l2_count = len(l1['children'])
            l2_width = (l1_width - (l2_gap_x * (l2_count - 1))) / l2_count
            
            # L1 밑에 L2 시작 (기본 간격 A 적용)
            current_y_for_l2 = y_l1 + h_l1 + v_gap_a

            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap_x))
                y_l2 = current_y_for_l2
                h_l2 = 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})

                # 재귀적으로 하위 노드 배치
                def stack_recursive(parent_node, px, py, pw, ph):
                    nonlocal layout_data
                    last_y = py + ph
                    
                    for idx, child in enumerate(parent_node['children']):
                        lvl = child['level']
                        # 1. 기본적으로 모든 수직 이동은 간격 A
                        gap = v_gap_a
                        
                        # 2. 형제 그룹이 시작될 때만(idx > 0) 레벨별 추가 여백 적용
                        if idx > 0:
                            gap += extra_gaps.get(lvl, extra_gaps[5] if lvl > 5 else 0)
                        
                        target_y = last_y + gap
                        
                        reduction = 0.3 * (lvl - 2)
                        c_w = max(l2_width - reduction, 2.0)
                        c_x = (px + pw) - c_w
                        c_h = 0.8
                        
                        layout_data.append({
                            'node': child, 'x': c_x, 'y': target_y, 'w': c_w, 'h': c_h, 'level': lvl
                        })
                        
                        # 자식 그룹 끝점 추적
                        last_y = stack_recursive(child, c_x, target_y, c_w, c_h)
                    
                    return last_y

                draw_end_y = stack_recursive(l2, x_l2, y_l2, l2_width, h_l2)
                    
    return layout_data

# --- 3. 미리보기 (Matplotlib) ---
def draw_preview(layout_data):
    fig, ax = plt.subplots(figsize=(10, 5.6))
    ax.set_xlim(0, 33.8)
    ax.set_ylim(0, 19.05)
    ax.invert_yaxis()
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, linewidth=1, edgecolor='#cccccc', facecolor='#ffffff'))
    for item in layout_data:
        lvl = item['level']
        color = '#1f497d' if lvl == 1 else '#365f91' if lvl == 2 else '#f2f2f2'
        edge = 'white' if lvl <= 2 else '#cccccc'
        rect = patches.Rectangle((item['x'], item['y']), item['w'], item['h'], linewidth=0.5, edgecolor=edge, facecolor=color)
        ax.add_patch(rect)
        display_text = item['node']['text'][:15]
        txt_color = 'white' if lvl <= 2 else 'black'
        ax.text(item['x'] + item['w']/2, item['y'] + item['h']/2, display_text, color=txt_color, fontsize=5.5, ha='center', va='center')
    ax.set_axis_off()
    st.pyplot(fig)

# --- 4. PPT 생성 (이전과 동일) ---
def generate_ppt(layout_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.8), Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for item in layout_data:
        shp = slide.shapes.add_shape(1, Cm(item['x']), Cm(item['y']), Cm(item['w']), Cm(item['h']))
        lvl = item['level']
        shp.fill.solid()
        if lvl == 1:
            shp.fill.fore_color.rgb = RGBColor(31, 73, 125)
            f_size, f_bold, f_color, align = Pt(12), True, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        elif lvl == 2:
            shp.fill.fore_color.rgb = RGBColor(54, 95, 145)
            f_size, f_bold, f_color, align = Pt(10), False, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        else:
            c = min(220 + (lvl * 5), 250)
            shp.fill.fore_color.rgb = RGBColor(c, c, c+5)
            shp.line.color.rgb = RGBColor(180, 180, 180)
            f_size, f_bold, f_color, align = Pt(8.5), False, RGBColor(0, 0, 0), PP_ALIGN.LEFT
        tf = shp.text_frame
        tf.text = item['node']['text']
        p = tf.paragraphs[0]
        p.font.size, p.font.bold, p.font.color.rgb, p.alignment = f_size, f_bold, f_color, align
    return prs

# --- 5. Streamlit UI ---
st.set_page_config(page_title="WBS Ultimate Designer", layout="wide")
st.sidebar.title("🎨 WBS 프로 디자인 설정")

with st.sidebar.expander("📏 캔버스 크기 (cm)", expanded=True):
    wbs_w = st.number_input("WBS 전체 너비", 10.0, 32.0, 31.0)
    wbs_h = st.number_input("WBS 전체 높이", 5.0, 18.0, 16.0)

with st.sidebar.expander("↕️ 상하 간격 정밀 설정 (cm)", expanded=True):
    v_gap_a = st.number_input("기준 수직 간격 (A)", 0.0, 5.0, 0.4, 0.05)
    st.divider()
    extra_l3 = st.number_input("L3 그룹 간 추가 여백", 0.0, 5.0, 0.3, 0.05)
    extra_l4 = st.number_input("L4 그룹 간 추가 여백", 0.0, 5.0, 0.2, 0.05)
    extra_l5 = st.number_input("L5+ 그룹 간 추가 여백", 0.0, 5.0, 0.1, 0.05)
    st.caption("※ 줄기(Group)가 바뀔 때만 추가 여백이 적용됩니다.")

with st.sidebar.expander("↔️ 좌우 간격 설정 (cm)", expanded=True):
    l1_gap_x = st.number_input("L1 좌우 간격", 0.0, 10.0, 1.2)
    l2_gap_x = st.number_input("L2 좌우 간격", 0.0, 5.0, 0.4)

config = {
    'wbs_w': wbs_w, 'wbs_h': wbs_h, 'l1_gap_x': l1_gap_x, 'l2_gap_x': l2_gap_x,
    'v_gap_a': v_gap_a, 'extra_l3': extra_l3, 'extra_l4': extra_l4, 'extra_l5': extra_l5
}

st.title("📊 WBS 마스터 디자이너 (통합 간격 제어)")
st.info("💡 기준 간격(A)은 모든 수직 리듬을 결정하며, 그룹 간 여백은 줄기 사이를 벌려 가독성을 높입니다.")

uploaded_file = st.file_uploader("파일 업로드", type=["xlsx", "pptx"])

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
        layout_data = calculate_layout(tree, config)
        st.subheader("🖼️ 실시간 디자인 미리보기")
        draw_preview(layout_data)
        if st.button("🚀 PPT 생성 및 다운로드", use_container_width=True):
            final_ppt = generate_ppt(layout_data)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 PPT 파일 다운로드", ppt_io, "Smart_WBS_Final.pptx")

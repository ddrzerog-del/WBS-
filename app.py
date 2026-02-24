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

# --- 2. 좌표 계산 로직 ---
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w = config['wbs_w']
    wbs_h = config['wbs_h']
    l1_gap_x = config['l1_gap_x']
    l2_gap_x = config['l2_gap_x']
    
    # 수직 간격 설정
    v_gaps = {
        2: config['v_gap_1_2'],
        3: config['v_gap_2_3'],
        4: config['v_gap_3_4'],
        'deep': config['v_gap_deep']
    }
    
    # 슬라이드 중앙 정렬 원점 (16:9 슬라이드 너비 33.8cm 기준)
    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2

    l1_count = len(root_nodes)
    if l1_count == 0: return []
    
    l1_width = (wbs_w - (l1_gap_x * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        x_l1 = start_x + (i * (l1_width + l1_gap_x))
        y_l1 = start_y
        h_l1 = 1.2
        layout_data.append({'node': l1, 'x': x_l1, 'y': y_l1, 'w': l1_width, 'h': h_l1, 'level': 1})

        if l1['children']:
            l2_count = len(l1['children'])
            l2_width = (l1_width - (l2_gap_x * (l2_count - 1))) / l2_count
            
            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap_x))
                y_l2 = y_l1 + h_l1 + v_gaps[2]
                h_l2 = 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})

                descendants = []
                get_all_descendants(l2, descendants)
                curr_y = y_l2 + h_l2
                
                for desc in descendants:
                    lvl = desc['level']
                    target_gap = v_gaps[3] if lvl == 3 else v_gaps[4] if lvl == 4 else v_gaps['deep']
                    curr_y += target_gap
                    
                    reduction = 0.4 * (lvl - 2)
                    d_w = max(l2_width - reduction, 2.0)
                    d_x = (x_l2 + l2_width) - d_w
                    d_h = 0.8
                    
                    layout_data.append({'node': desc, 'x': d_x, 'y': curr_y, 'w': d_w, 'h': d_h, 'level': lvl})
                    curr_y += d_h
                    
    return layout_data

# --- 3. 미리보기 (Matplotlib) ---
def draw_preview(layout_data):
    fig, ax = plt.subplots(figsize=(10, 5.6))
    ax.set_xlim(0, 33.8)
    ax.set_ylim(0, 19.05)
    ax.invert_yaxis()
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, linewidth=1, edgecolor='black', facecolor='#f9f9f9', alpha=0.5))

    for item in layout_data:
        lvl = item['level']
        color = '#1f497d' if lvl == 1 else '#365f91' if lvl == 2 else '#d9d9d9'
        rect = patches.Rectangle((item['x'], item['y']), item['w'], item['h'], linewidth=1, edgecolor='white', facecolor=color)
        ax.add_patch(rect)
        
        display_text = item['node']['text'][:10]
        txt_color = 'white' if lvl <= 2 else 'black'
        ax.text(item['x'] + item['w']/2, item['y'] + item['h']/2, display_text, color=txt_color, fontsize=6, ha='center', va='center')

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
        shp.fill.solid()
        if lvl == 1:
            shp.fill.fore_color.rgb = RGBColor(31, 73, 125)
            f_size, f_bold, f_color, align = Pt(12), True, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        elif lvl == 2:
            shp.fill.fore_color.rgb = RGBColor(54, 95, 145)
            f_size, f_bold, f_color, align = Pt(10), False, RGBColor(255, 255, 255), PP_ALIGN.CENTER
        else:
            c = min(210 + (lvl * 8), 250)
            shp.fill.fore_color.rgb = RGBColor(c, c, c+5)
            shp.line.color.rgb = RGBColor(180, 180, 180)
            f_size, f_bold, f_color, align = Pt(9), False, RGBColor(0, 0, 0), PP_ALIGN.LEFT
        tf = shp.text_frame
        tf.text = item['node']['text']
        p = tf.paragraphs[0]
        p.font.size, p.font.bold, p.font.color.rgb, p.alignment = f_size, f_bold, f_color, align
    return prs

# --- 5. Streamlit UI ---
st.set_page_config(page_title="WBS Designer Pro", layout="wide")

st.sidebar.title("🎨 WBS 상세 디자인")

with st.sidebar.expander("📏 전체 영역 (cm)", expanded=True):
    wbs_w = st.number_input("가로 너비", 10.0, 32.0, 30.0, 0.5)
    wbs_h = st.number_input("세로 높이", 5.0, 18.0, 15.0, 0.5)

with st.sidebar.expander("↔️ 좌우 간격 (cm)", expanded=True):
    l1_gap_x = st.number_input("대그룹(L1) 간격", 0.0, 10.0, 1.5, 0.1)
    l2_gap_x = st.number_input("소그룹(L2) 간격", 0.0, 5.0, 0.5, 0.1)

with st.sidebar.expander("↕️ 수직 간격 설정 (황금비율)", expanded=True):
    auto_golden = st.checkbox("황금비율($\phi$) 모드 사용", value=True)
    base_v_gap = st.number_input("기준 간격 (L1→L2)", 0.1, 5.0, 0.8, 0.1)

    if auto_golden:
        v_gap_1_2 = base_v_gap
        v_gap_2_3 = round(v_gap_1_2 * 0.618, 2)
        v_gap_3_4 = round(v_gap_2_3 * 0.618, 2)
        v_gap_deep = round(v_gap_3_4 * 0.618, 2)
        st.caption(f"자동계산: {v_gap_1_2} → {v_gap_2_3} → {v_gap_3_4} → {v_gap_deep}")
    else:
        v_gap_1_2 = st.number_input("1→2 레벨 간격", 0.0, 5.0, 0.6, 0.05)
        v_gap_2_3 = st.number_input("2→3 레벨 간격", 0.0, 5.0, 0.4, 0.05)
        v_gap_3_4 = st.number_input("3→4 레벨 간격", 0.0, 5.0, 0.2, 0.05)
        v_gap_deep = st.number_input("4레벨 이상 간격", 0.0, 5.0, 0.1, 0.05)

config = {
    'wbs_w': wbs_w, 'wbs_h': wbs_h, 
    'l1_gap_x': l1_gap_x, 'l2_gap_x': l2_gap_x,
    'v_gap_1_2': v_gap_1_2, 'v_gap_2_3': v_gap_2_3, 
    'v_gap_3_4': v_gap_3_4, 'v_gap_deep': v_gap_deep
}

st.title("📊 WBS 프로 디자이너 (황금비율 버전)")

uploaded_file = st.file_uploader("엑셀 또는 PPT 파일 업로드", type=["xlsx", "pptx"])

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
        
        st.subheader("🖼️ 슬라이드 미리보기")
        draw_preview(layout_data)
        
        if st.button("🚀 최종 PPT 생성", use_container_width=True):
            final_ppt = generate_ppt(layout_data)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 PPT 다운로드", ppt_io, "Smart_WBS_Golden.pptx")

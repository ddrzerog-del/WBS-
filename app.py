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

# --- 2. 좌표 계산 로직 (그룹/아이템 간격 분리) ---
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w = config['wbs_w']
    wbs_h = config['wbs_h']
    l1_gap_x = config['l1_gap_x']
    l2_gap_x = config['l2_gap_x']
    
    item_v_gap = config['item_v_gap']   # 아이템 간 기본 간격
    group_v_gap = config['group_v_gap'] # 그룹(줄기) 간 추가 여백

    # 슬라이드 중앙 정렬 원점
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
            
            # 1레벨 아래 2레벨은 첫 시작이므로 그룹 간격 적용
            current_y_for_l2 = y_l1 + h_l1 + group_v_gap

            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap_x))
                y_l2 = current_y_for_l2
                h_l2 = 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})

                # 재귀적으로 하위 노드 배치
                def draw_recursive(parent_node, px, py, pw, ph):
                    nonlocal layout_data
                    last_y = py + ph
                    
                    for idx, child in enumerate(parent_node['children']):
                        # 첫 자식은 부모에게 밀착(item_v_gap)
                        # 두 번째 형제부터는 '그룹 간 여백(group_v_gap)'을 추가로 적용
                        gap = item_v_gap if idx == 0 else (item_v_gap + group_v_gap)
                        
                        target_y = last_y + gap
                        
                        # 너비 축소 및 우측 정렬
                        reduction = 0.3 * (child['level'] - 2)
                        c_w = max(l2_width - reduction, 2.0)
                        c_x = (px + pw) - c_w
                        c_h = 0.8
                        
                        layout_data.append({
                            'node': child, 'x': c_x, 'y': target_y, 'w': c_w, 'h': c_h, 'level': child['level']
                        })
                        
                        # 자식의 가족들 전체 높이를 계산하여 반환
                        last_y = draw_recursive(child, c_x, target_y, c_w, c_h)
                    
                    return last_y

                draw_recursive(l2, x_l2, y_l2, l2_width, h_l2)
                    
    return layout_data

# --- 3. 미리보기 (Matplotlib) ---
def draw_preview(layout_data):
    fig, ax = plt.subplots(figsize=(10, 5.6))
    ax.set_xlim(0, 33.8)
    ax.set_ylim(0, 19.05)
    ax.invert_yaxis()
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, linewidth=1, edgecolor='gray', facecolor='#fdfdfd', alpha=0.5))
    for item in layout_data:
        lvl = item['level']
        color = '#1f497d' if lvl == 1 else '#365f91' if lvl == 2 else '#efefef'
        rect = patches.Rectangle((item['x'], item['y']), item['w'], item['h'], linewidth=0.5, edgecolor='#aaaaaa', facecolor=color)
        ax.add_patch(rect)
        display_text = item['node']['text'][:12]
        txt_color = 'white' if lvl <= 2 else 'black'
        ax.text(item['x'] + item['w']/2, item['y'] + item['h']/2, display_text, color=txt_color, fontsize=6, ha='center', va='center')
    ax.set_axis_off()
    st.pyplot(fig)

# --- 4. PPT 생성 ---
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
            shp.line.color.rgb = RGBColor(200, 200, 200)
            f_size, f_bold, f_color, align = Pt(9), False, RGBColor(0, 0, 0), PP_ALIGN.LEFT
        tf = shp.text_frame
        tf.text = item['node']['text']
        p = tf.paragraphs[0]
        p.font.size, p.font.bold, p.font.color.rgb, p.alignment = f_size, f_bold, f_color, align
    return prs

# --- 5. Streamlit UI ---
st.set_page_config(page_title="WBS Master Designer", layout="wide")

st.sidebar.title("🎨 WBS 레이아웃 커스텀")

with st.sidebar.expander("📏 전체 캔버스 (cm)", expanded=True):
    wbs_w = st.number_input("WBS 전체 너비", 10.0, 32.0, 30.0, 0.5)
    wbs_h = st.number_input("WBS 전체 높이", 5.0, 18.0, 15.0, 0.5)

with st.sidebar.expander("↔️ 좌우 간격 (cm)", expanded=True):
    l1_gap_x = st.number_input("대그룹(L1) 간격", 0.0, 10.0, 1.5, 0.1)
    l2_gap_x = st.number_input("소그룹(L2) 간격", 0.0, 5.0, 0.5, 0.1)

with st.sidebar.expander("↕️ 상하 간격 정밀 조정 (cm)", expanded=True):
    item_v_gap = st.number_input("아이템 간 간격 (기본)", 0.0, 2.0, 0.1, 0.05)
    group_v_gap = st.number_input("그룹(줄기) 간 추가 여백", 0.0, 5.0, 0.8, 0.1)
    st.caption("그룹 간 여백을 키우면 다른 줄기 사이가 벌어집니다.")

config = {
    'wbs_w': wbs_w, 'wbs_h': wbs_h, 
    'l1_gap_x': l1_gap_x, 'l2_gap_x': l2_gap_x,
    'item_v_gap': item_v_gap, 'group_v_gap': group_v_gap
}

st.title("📊 WBS 마스터 디자이너")
st.info("아이템 간 간격은 '촘촘하게', 그룹 간 추가 여백은 '넓게' 설정하면 최고의 가독성이 나옵니다.")

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
        
        st.subheader("🖼️ 디자인 미리보기")
        draw_preview(layout_data)
        
        if st.button("🚀 설정대로 PPT 생성", use_container_width=True):
            final_ppt = generate_ppt(layout_data)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 PPT 파일 다운로드", ppt_io, "Customized_WBS.pptx")

# --- 사이드바 설정 부분 수정 ---
with st.sidebar.expander("↕️ 수직 간격 설정 (황금비율)", expanded=True):
    auto_golden = st.checkbox("황금비율($\phi$) 모드 사용", value=True)
    base_v_gap = st.number_input("기준 간격 (L1→L2, cm)", 0.1, 5.0, 1.0, 0.1)

    if auto_golden:
        # 황금비율 자동 계산
        v_gap_1_2 = base_v_gap
        v_gap_2_3 = round(v_gap_1_2 * 0.618, 2)
        v_gap_3_4 = round(v_gap_2_3 * 0.618, 2)
        v_gap_deep = round(v_gap_3_4 * 0.618, 2)
        
        st.caption(f"자동 설정됨: L2→L3({v_gap_2_3}cm), L3→L4({v_gap_3_4}cm)...")
    else:
        # 수동 설정
        v_gap_1_2 = st.number_input("1레벨 → 2레벨", 0.0, 5.0, 0.6, 0.05)
        v_gap_2_3 = st.number_input("2레벨 → 3레벨", 0.0, 5.0, 0.4, 0.05)
        v_gap_3_4 = st.number_input("3레벨 → 4레벨", 0.0, 5.0, 0.2, 0.05)
        v_gap_deep = st.number_input("4레벨 이상 간격", 0.0, 5.0, 0.1, 0.05)

# --- 실제 계산 로직에 적용 ---
config = {
    'v_gap_1_2': v_gap_1_2,
    'v_gap_2_3': v_gap_2_3,
    'v_gap_3_4': v_gap_3_4,
    'v_gap_deep': v_gap_deep,
    # ... 기타 설정 ...
}

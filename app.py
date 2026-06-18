import streamlit as st
import pandas as pd
import io

# 모바일 화면에 맞게 꽉 차는 레이아웃 설정
st.set_page_config(page_title="GRE PIPE 모바일 체크시트", layout="wide", initial_sidebar_state="collapsed")

# 데이터 누적을 위한 세션 상태(Session State) 초기화
if 'align_data' not in st.session_state:
    st.session_state['align_data'] = []
if 'torque_data' not in st.session_state:
    st.session_state['torque_data'] = []

st.title("🛠️ GRE PIPE 모바일 체크시트")
st.markdown("---")

# ==========================================
# 1. 기본 확인 사항 섹션
# ==========================================
st.header("1. 기본 점검 사항")
questions = [
    "파이프 작업 전 도면은 확인 하였는가?",
    "파이프 조인트부 선수, 선미에 각각 1개씩 SUPPORT는 설치되어있는가?",
    "파이프 내부에 이물질은 없는가?",
    "파이프 외부에 이물질이나 데미지는 없는가?",
    "파이프 설치 시 alignment 는 허용값안에 들어 오는가?",
    "파이프 설치 시 기준에 맞는 볼트 / 너트 / 와셔를 사용 하였는가?",
    "Earth Wire는 정확한 위치에 설치 하였는가?",
    "토크작업 전 토크렌치 교정검정일은 확인 하였는가? (교정성적서)",
    "볼트 체결시 메이커 매뉴얼 순서에 맞게 작업 하였는가?",
    "배관 커버링은 제대로 설치되었는가? (함석 + 난연성커버)"
]

# 기본 점검 결과를 저장할 딕셔너리
basic_results = {}

for i, q in enumerate(questions, 1):
    st.markdown(f"**Q{i}. {q}**")
    col1, col2 = st.columns([1, 1])
    with col1:
        status = st.radio(
            "점검결과", 
            ["양호", "불량", "해당없음"], 
            key=f"status_{i}", 
            horizontal=True, 
            label_visibility="collapsed"
        )
    with col2:
        remark = st.text_input(
            "조치내용", 
            placeholder="조치내용 입력...", 
            key=f"remark_{i}", 
            label_visibility="collapsed"
        )
    basic_results[f"Q{i}"] = {"점검항목": q, "결과": status, "조치내용": remark}
    st.write("") 

st.markdown("<h4 style='color: red; text-align: center;'>※ 단 LEAK 발생시 담당PRO와 협의 후 MAXIMUM TORQUE값 사용가능</h4>", unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# 2. ALIGNMENT CHECK RESULT 섹션 (영/한 혼용)
# ==========================================
st.header("2. ALIGNMENT CHECK")
with st.expander("👉 ALIGNMENT 데이터 입력하기 (터치하여 열기)"):
    col1, col2 = st.columns(2)
    with col1:
        dia_align = st.text_input("DIA (관경)", key="dia_a")
        top = st.text_input("TOP (상부)", key="top")
        port = st.text_input("PORT (좌현)", key="port")
        gap = st.text_input("GAP (간극)", key="gap")
    with col2:
        coupling_no = st.text_input("COUPLING NUMBER (커플링 번호)", key="coup_no")
        bottom = st.text_input("BOTTOM (하부)", key="bottom")
        stb = st.text_input("STB (우현)", key="stb")
        remark_align = st.text_input("REMARK (비고)", key="rem_a")
    
    if st.button("➕ ALIGNMENT 데이터 추가", use_container_width=True):
        st.session_state['align_data'].append({
            "DIA (관경)": dia_align,
            "COUPLING NUMBER (커플링 번호)": coupling_no,
            "TOP (상부)": top, 
            "BOTTOM (하부)": bottom,
            "PORT (좌현)": port, 
            "STB (우현)": stb,
            "GAP (간극)": gap, 
            "REMARK (비고)": remark_align
        })
        st.success("데이터가 아래 리스트에 추가되었습니다.")

# 누적된 Alignment 데이터 테이블 출력
if st.session_state['align_data']:
    st.dataframe(pd.DataFrame(st.session_state['align_data']), use_container_width=True)

st.markdown("---")

# ==========================================
# 3. TORQUE VALUE CHECK 섹션 (영/한 혼용)
# ==========================================
st.header("3. TORQUE VALUE CHECK")
with st.expander("👉 TORQUE 데이터 입력하기 (터치하여 열기)"):
    col1, col2 = st.columns(2)
    with col1:
        dia_torque = st.text_input("DIA (관경)", key="dia_t")
        elem_no_2 = st.text_input("ELEM.NO (도면/요소번호 2)", key="elem2")
        torque_serial = st.text_input("TORQUE WRENCH S/N (시리얼 넘버)", key="serial")
    with col2:
        elem_no_1 = st.text_input("ELEM.NO (도면/요소번호 1)", key="elem1")
        torque_val = st.text_input("TORQUE VALUE (토크 값)", key="t_val")
        
    if st.button("➕ TORQUE 데이터 추가", use_container_width=True):
        st.session_state['torque_data'].append({
            "DIA (관경)": dia_torque,
            "ELEM.NO (도면/요소번호 1)": elem_no_1,
            "ELEM.NO (도면/요소번호 2)": elem_no_2,
            "TORQUE VALUE (토크 값)": torque_val,
            "TORQUE WRENCH S/N (시리얼 넘버)": torque_serial
        })
        st.success("데이터가 아래 리스트에 추가되었습니다.")

# 누적된 Torque 데이터 테이블 출력
if st.session_state['torque_data']:
    st.dataframe(pd.DataFrame(st.session_state['torque_data']), use_container_width=True)

st.markdown("---")

# ==========================================
# 4. 현장 증빙 사진
# ==========================================
st.header("4. 현장 증빙 사진")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**📸 토크렌치 교정번호**")
    cam_photo1 = st.camera_input("카메라 촬영", key="cam1", label_visibility="collapsed")
    upload_photo1 = st.file_uploader("또는 업로드", type=["png", "jpg", "jpeg"], key="up1")

with col2:
    st.markdown("**📸 토크렌치 세팅 값**")
    cam_photo2 = st.camera_input("카메라 촬영", key="cam2", label_visibility="collapsed")
    upload_photo2 = st.file_uploader("또는 업로드", type=["png", "jpg", "jpeg"], key="up2")

st.markdown("---")

# ==========================================
# 5. 엑셀 변환 및 저장 로직
# ==========================================
def generate_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. 기본 점검 사항 시트
        df_basic = pd.DataFrame(basic_results).T
        df_basic.to_excel(writer, sheet_name='기본점검사항', index=False)
        
        # 2. Alignment 데이터 시트
        if st.session_state['align_data']:
            df_align = pd.DataFrame(st.session_state['align_data'])
        else:
            df_align = pd.DataFrame(columns=[
                "DIA (관경)", "COUPLING NUMBER (커플링 번호)", "TOP (상부)", 
                "BOTTOM (하부)", "PORT (좌현)", "STB (우현)", "GAP (간극)", "REMARK (비고)"
            ])
        df_align.to_excel(writer, sheet_name='ALIGNMENT_체크', index=False)
        
        # 3. Torque 데이터 시트
        if st.session_state['torque_data']:
            df_torque = pd.DataFrame(st.session_state['torque_data'])
        else:
            df_torque = pd.DataFrame(columns=[
                "DIA (관경)", "ELEM.NO (도면/요소번호 1)", "ELEM.NO (도면/요소번호 2)", 
                "TORQUE VALUE (토크 값)", "TORQUE WRENCH S/N (시리얼 넘버)"
            ])
        df_torque.to_excel(writer, sheet_name='TORQUE_체크', index=False)

    return output.getvalue()

# 하단 제어 버튼 (임시저장 및 다운로드)
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("💾 임시저장", use_container_width=True):
        st.toast("현재까지 작성된 내용이 브라우저 세션에 임시저장 되었습니다.")

with col_btn2:
    excel_data = generate_excel()
    st.download_button(
        label="✅ 점검결과 제출 (Excel 다운로드)",
        data=excel_data,
        file_name="GRE_PIPE_점검시트.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
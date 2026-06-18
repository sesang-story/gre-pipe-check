import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as xlImage
from openpyxl.utils import get_column_letter
from datetime import datetime

# 구글 API 라이브러리
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# ⚙️ 구글 클라우드 고유 ID 설정
# ==========================================
SPREADSHEET_ID = "1GNKbHoS7950PqjZNB0xqJIRBuEQnqSGbpCqpOiAGI-U"

# 모바일 뷰 최적화
st.set_page_config(page_title="GRE PIPE 모바일 체크시트", layout="wide", initial_sidebar_state="collapsed")

# 세션 상태 초기화
if 'align_data' not in st.session_state: st.session_state['align_data'] = []
if 'torque_data' not in st.session_state: st.session_state['torque_data'] = []
if 'photo1' not in st.session_state: st.session_state['photo1'] = None
if 'photo2' not in st.session_state: st.session_state['photo2'] = None

st.title("🛠️ GRE PIPE 모바일 체크시트")
st.markdown("---")

# ==========================================
# 📝 보고서 현장 정보 (자동 조합 로직 추가)
# ==========================================
with st.expander("📄 보고서 현장 정보 (보고서 및 클라우드 동기화)", expanded=True):
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        doc_date = st.date_input("작성일자", datetime.now())
        hull_no = st.text_input("호선 (Hull No.)", placeholder="예: SN2717")
        block_no = st.text_input("블록 (Block)", placeholder="예: D620S")
        tag_no = st.text_input("TAG NO.", placeholder="예: CG001")
        
    with col_m2:
        dept_name = st.selectbox("소속", ["선택 안함", "1직1반", "1직2반", "2직1반", "3직1반", "3직2반", "4직2반", "4직3반", "대영기업"])
        doc_area = st.text_input("AREA", placeholder="예: E21 현장")
        doc_author = st.text_input("이름 (작성자)", value="")

    # 💡 문서번호 자동 생성 로직
    auto_doc_no = "SHI-GRE"
    if hull_no: auto_doc_no += f"-{hull_no.strip()}"
    if block_no: auto_doc_no += f"-{block_no.strip()}"
    if tag_no: auto_doc_no += f"-{tag_no.strip()}"
    
    # 아무것도 입력 안 했을 때의 기본값 보호
    if auto_doc_no == "SHI-GRE":
        auto_doc_no = f"SHI-GRE-{datetime.now().strftime('%Y%m')}-001"
        
    doc_no = auto_doc_no
    st.text_input("문서번호 (자동생성)", value=doc_no, disabled=True)

# 소속 선택 안 했을 경우 빈칸 처리
final_dept_name = "" if dept_name == "선택 안함" else dept_name

st.markdown("---")

# ==========================================
# 1. 기본 확인 사항
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

basic_results = []
for i, q in enumerate(questions, 1):
    st.markdown(f"**Q{i}. {q}**")
    col1, col2 = st.columns([1, 1])
    with col1:
        status = st.radio("점검결과", ["양호", "불량", "해당없음"], key=f"status_{i}", horizontal=True, label_visibility="collapsed")
    with col2:
        remark = st.text_input("조치내용", placeholder="조치내용 입력...", key=f"remark_{i}", label_visibility="collapsed")
    basic_results.append({"순번": i, "항목": q, "결과": status, "비고": remark if remark else "-"})
    st.write("") 

st.markdown("---")

# ==========================================
# 2. 데이터 입력
# ==========================================
st.header("2. 데이터 입력")
st.info("💡 값을 입력한 후 [➕ 추가] 버튼을 안 눌러도, 입력된 값은 보고서에 자동으로 1줄 반영됩니다!")
tab1, tab2 = st.tabs(["ALIGNMENT CHECK", "TORQUE CHECK"])

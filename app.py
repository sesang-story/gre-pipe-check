import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as xlImage
from openpyxl.utils import get_column_letter
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# ⚙️ 구글 클라우드 고유 ID 설정
# ==========================================
SPREADSHEET_ID = "1GNKbHoS7950PqjZNB0xqJIRBuEQnqSGbpCqpOiAGI-U"

st.set_page_config(page_title="GRE PIPE 모바일 체크시트", layout="wide", initial_sidebar_state="collapsed")

if 'align_data' not in st.session_state: st.session_state['align_data'] = []
if 'torque_data' not in st.session_state: st.session_state['torque_data'] = []
if 'photo1' not in st.session_state: st.session_state['photo1'] = None
if 'photo2' not in st.session_state: st.session_state['photo2'] = None

st.title("🛠️ GRE PIPE 모바일 체크시트")
st.markdown("---")

# ==========================================
# 📝 1. 보고서 현장 정보
# ==========================================
with st.expander("📄 보고서 현장 정보 (보고서 및 클라우드 동기화)", expanded=True):
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        doc_date = st.date_input("작성일자", datetime.now())
        hull_no = st.text_input("호선 (Hull No.)", placeholder="예: SN2717")
        block_no = st.text_input("블록 (Block)", placeholder="예: D620S")
        tag_no = st.text_input("TAG NO.", placeholder="예: CG001")
        
    with col_m2:
        dept_name = st.selectbox("소속", ["선택 안함", "선행의장 1과", "1직1반", "1직2반", "2직1반", "3직1반", "3직2반", "4직2반", "4직3반", "대영기업"])
        doc_area = st.text_input("AREA", placeholder="예: 3001")
        doc_author = st.text_input("이름 (작성자)", value="")

    parts = ["SHI-GRE"]
    if hull_no: parts.append(hull_no.strip())
    if block_no: parts.append(block_no.strip())
    if tag_no: parts.append(tag_no.strip())
    
    if len(parts) == 1: doc_no = f"SHI-GRE-{datetime.now().strftime('%Y%m')}-001"
    else: doc_no = "-".join(parts)
        
    st.info(f"**📑 문서번호 (자동생성):** {doc_no}")

final_dept_name = "" if dept_name == "선택 안함" else dept_name

st.markdown("---")

# ==========================================
# 2. 기본 점검 사항 (데이터 취합)
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

# 전체 판정 로직: 10개 중 하나라도 '양호'가 아니면 '불량'
overall_status = "양호"
for item in basic_results:
    if item["결과"] != "양호":
        overall_status = "불량"
        break

st.markdown("---")

# ==========================================
# 3. 데이터 입력 (얼라인먼트 & 토크)
# ==========================================
st.header("2. 데이터 입력")
st.info("💡 값을 입력한 후 [➕ 추가] 버튼을 안 눌러도, 입력된 값은 보고서에 자동으로 1줄 반영됩니다!")

dia_a = top = port = gap = coup_no = bottom = stb = ""
rem_a = "정상 범위 내"
dia_t = elem2 = elem1 = t_val = ""
serial = "P2024500178"

tab1, tab2 = st.tabs(["ALIGNMENT CHECK", "TORQUE CHECK"])

with tab1:
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        dia_a = st.text_input("DIA (관경)", key="dia_a")
        top = st.text_input("TOP (상부)", key="top")
        port = st.text_input("PORT (좌현)", key="port")
        gap = st.text_input("GAP (간극)", key="gap")
    with col_a2:
        coup_no = st.text_input("COUPLING NUMBER", key="coup_no")
        bottom = st.text_input("BOTTOM (하부)", key="bottom")
        stb = st.text_input("STB (우현)", key="stb")
        rem_a = st.text_input("REMARK (비고)", key="rem_a", value="정상 범위 내")
    if st.button("➕ ALIGNMENT 리스트에 추가", use_container_width=True):
        st.session_state['align_data'].append([dia_a, coup_no, top, bottom, port, stb, gap, rem_a])
        st.success("추가되었습니다.")

with tab2:
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        dia_t = st.text_input("DIA (관경)", key="dia_t")
        elem2 = st.text_input("ELEM.NO (2)", key="elem2")
        serial = st.text_input("토크렌치 S/N", key="serial", value="P2024500178")
    with col_t2:
        elem1 = st.text_input("ELEM.NO (1)", key="elem1")
        t_val = st.text_input("TORQUE VALUE", key="t_val")
    if st.button("➕ TORQUE 리스트에 추가", use_container_width=True):
        st.session_state['torque_data'].append([dia_t, elem1, elem2, t_val, serial])
        st.success("추가되었습니다.")

st.markdown("---")

# ==========================================
# 4. 현장 증빙 사진 
# ==========================================
st.header("3. 현장 증빙 사진")
col_p1, col_p2 = st.columns(2)

with col_p1:
    st.markdown("**📸 토크렌치 교정번호**")
    cam1 = st.camera_input("카메라 1", key="cam1", label_visibility="collapsed")
    up1 = st.file_uploader("업로드 1", type=["png", "jpg", "jpeg"], key="up1", label_visibility="collapsed")
    if cam1: st.session_state['photo1'] = cam1.getvalue()
    elif up1: st.session_state['photo1'] = up1.getvalue()

with col_p2:
    st.markdown("**📸 토크렌치 세팅 값**")
    cam2 = st.camera_input("카메라 2", key="cam2", label_visibility="collapsed")
    up2 = st.file_uploader("업로드 2", type=["png", "jpg", "jpeg"], key="up2", label_visibility="collapsed")
    if cam2: st.session_state['photo2'] = cam2.getvalue()
    elif up2: st.session_state['photo2'] = up2.getvalue()

st.markdown("---")

final_align_data = st.session_state['align_data'].copy()
if not final_align_data and (dia_a or coup_no): 
    final_align_data = [[dia_a, coup_no, top, bottom, port, stb, gap, rem_a]]

final_torque_data = st.session_state['torque_data'].copy()
if not final_torque_data and (dia_t or elem1 or elem2):
    final_torque_data = [[dia_t, elem1, elem2, t_val, serial]]

# ==========================================
# 5. 엑셀 서식 생성
# ==========================================
def generate_report(align_list, torque_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "품질검사보고서"
    ws.views.sheetView[0].showGridLines = False

    NAVY_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    SUB_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    WHITE_FONT = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    SUB_FONT = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")
    BODY_FONT = Font(name="맑은 고딕", size=10)
    BODY_BOLD = Font(name="맑은 고딕", size=10, bold=True)
    ALERT_FONT = Font(name="맑은 고딕", size=10, bold=True, color="FF0000")
    
    ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
    THIN_BORDER = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    def apply_style(cell, font, align, border, fill=None):
        cell.font = font
        cell.alignment = align
        cell.border = border
        if fill: cell.fill = fill

    ws.merge_cells("A1:H2")
    title = ws["A1"]
    title.value = "GRE PIPE INSTALLATION & TORQUE INSPECTION REPORT"
    apply_style(title, Font(name="맑은 고딕", size=16, bold=True, color="FFFFFF"), ALIGN_C, THIN_BORDER, NAVY_FILL)

    row_idx = 4
    meta_info = [
        ("작성일자", str(doc_date), "문서번호", doc_no),
        ("호선 (Hull No.)", hull_no, "블록 (Block)", block_no),
        ("TAG NO.", tag_no, "소속", final_dept_name),
        ("이름", doc_author, "AREA", doc_area)
    ]
    
    for lbl1, val1, lbl2, val2 in meta_info:
        ws.row_dimensions

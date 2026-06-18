import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as xlImage
from openpyxl.utils import get_column_letter
from datetime import datetime

# 구글 API 라이브러리
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# ⚙️ 구글 클라우드 고유 ID 설정 (프로님의 ID로 변경해 주세요)
# ==========================================
SPREADSHEET_ID = "https://docs.google.com/spreadsheets/d/1GNKbHoS7950PqjZNB0xqJIRBuEQnqSGbpCqpOiAGI-U/edit?gid=1848578634#gid=1848578634"
DRIVE_FOLDER_ID = "https://drive.google.com/drive/u/0/folders/1tjT9Tw8xCty-gjZkCT8_rleyQHYodoS1"

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

basic_results = []
for i, q in enumerate(questions, 1):
    st.markdown(f"**Q{i}. {q}**")
    col1, col2 = st.columns([1, 1])
    with col1:
        status = st.radio("점검결과", ["양호", "불량", "해당없음"], key=f"status_{i}", horizontal=True, label_visibility="collapsed")
    with col2:
        remark = st.text_input("조치내용", placeholder="조치내용...", key=f"remark_{i}", label_visibility="collapsed")
    basic_results.append({"순번": i, "항목": q, "결과": status, "비고": remark})
    st.write("") 

st.markdown("<h4 style='color: red; text-align: center;'>※ 단 LEAK 발생시 담당PRO와 협의 후 MAXIMUM TORQUE값 사용가능</h4>", unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# 2. 데이터 입력 섹션
# ==========================================
st.header("2. 데이터 입력")
tab1, tab2 = st.tabs(["ALIGNMENT CHECK", "TORQUE CHECK"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        dia_a = st.text_input("DIA (관경)", key="dia_a")
        top = st.text_input("TOP (상부)", key="top")
        port = st.text_input("PORT (좌현)", key="port")
        gap = st.text_input("GAP (간극)", key="gap")
    with col2:
        coup_no = st.text_input("COUPLING NUMBER", key="coup_no")
        bottom = st.text_input("BOTTOM (하부)", key="bottom")
        stb = st.text_input("STB (우현)", key="stb")
        rem_a = st.text_input("REMARK (비고)", key="rem_a")
    if st.button("➕ ALIGNMENT 추가", use_container_width=True):
        st.session_state['align_data'].append([dia_a, coup_no, top, bottom, port, stb, gap, rem_a])
        st.success("리스트에 추가되었습니다.")
    if st.session_state['align_data']:
        st.dataframe(pd.DataFrame(st.session_state['align_data'], columns=["DIA", "COUP_NO", "TOP", "BOTTOM", "PORT", "STB", "GAP", "REMARK"]), use_container_width=True)

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        dia_t = st.text_input("DIA (관경)", key="dia_t")
        elem2 = st.text_input("ELEM.NO (2)", key="elem2")
        serial = st.text_input("토크렌치 S/N", key="serial")
    with col2:
        elem1 = st.text_input("ELEM.NO (1)", key="elem1")
        t_val = st.text_input("TORQUE VALUE", key="t_val")
    if st.button("➕ TORQUE 추가", use_container_width=True):
        st.session_state['torque_data'].append([dia_t, elem1, elem2, t_val, serial, "", "", ""])
        st.success("리스트에 추가되었습니다.")
    if st.session_state['torque_data']:
        st.dataframe(pd.DataFrame(st.session_state['torque_data']).iloc[:, :5], use_container_width=True)

st.markdown("---")

# ==========================================
# 3. 현장 증빙 사진 섹션
# ==========================================
st.header("3. 현장 증빙 사진")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**📸 토크렌치 교정번호**")
    cam1 = st.camera_input("카메라 1", key="cam1", label_visibility="collapsed")
    up1 = st.file_uploader("업로드 1", type=["png", "jpg", "jpeg"], key="up1", label_visibility="collapsed")
    if cam1: st.session_state['photo1'] = cam1.getvalue()
    elif up1: st.session_state['photo1'] = up1.getvalue()

with col2:
    st.markdown("**📸 토크렌치 세팅 값**")
    cam2 = st.camera_input("카메라 2", key="cam2", label_visibility="collapsed")
    up2 = st.file_uploader("업로드 2", type=["png", "jpg", "jpeg"], key="up2", label_visibility="collapsed")
    if cam2: st.session_state['photo2'] = cam2.getvalue()
    elif up2: st.session_state['photo2'] = up2.getvalue()

st.markdown("---")

# ==========================================
# 4. 엑셀 보고서 양식 생성 함수
# ==========================================
def generate_report():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "품질검사보고서"
    ws.views.sheetView[0].showGridLines = False

    NAVY_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    WHITE_FONT = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    BODY_FONT = Font(name="맑은 고딕", size=10)
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
    title.font = Font(name="맑은 고딕", size=16, bold=True)
    title.alignment = ALIGN_C

    row_idx = 4
    ws.cell(row=row_idx, column=1, value="1. 기본 점검 사항").font = Font(name="맑은 고딕", size=12, bold=True)
    row_idx += 1
    
    headers_s1 = ["순번", "점검 항목 (Inspection Items)", "", "", "", "", "점검결과", "조치내용"]
    ws.merge_cells(f"B{row_idx}:F{row_idx}")
    for col, h in enumerate(headers_s1, 1):
        c = ws.cell(row=row_idx, column=col, value=h)
        apply_style(c, WHITE_FONT, ALIGN_C, THIN_BORDER, NAVY_FILL)
    
    for item in basic_results:
        row_idx += 1
        ws.row_dimensions[row_idx].height = 25
        ws.cell(row=row_idx, column=1, value=item["순번"])
        ws.merge_cells(f"B{row_idx}:F{row_idx}")
        ws.cell(row=row_idx, column=2, value=item["항목"])
        ws.cell(row=row_idx, column=7, value=item["결과"])
        ws.cell(row=row_idx, column=8, value=item["비고"])
        for col in range(1, 9): apply_style(ws.cell(row=row_idx, column=col), BODY_FONT, ALIGN_C if col!=2 else ALIGN_L, THIN_BORDER)
        
    row_idx += 2
    ws.cell(row=row_idx, column=1, value="2. ALIGNMENT & TORQUE CHECK RESULT").font = Font(name="맑은 고딕", size=12, bold=True)
    row_idx += 1
    
    align_h = ["DIA", "COUPLING NO", "TOP", "BOTTOM", "PORT", "STB", "GAP", "REMARK"]
    for col, h in enumerate(align_h, 1): apply_style(ws.cell(row=row_idx, column=col, value=h), WHITE_FONT, ALIGN_C, THIN_BORDER, NAVY_FILL)
    for row_data in st.session_state['align_data']:
        row_idx += 1
        for col, val in enumerate(row_data, 1): apply_style(ws.cell(row=row_idx, column=col, value=val), BODY_FONT, ALIGN_C, THIN_BORDER)

    row_idx += 2
    torque_h = ["DIA", "ELEM.NO 1", "ELEM.NO 2", "토크 값", "토크렌치 S/N", "", "", ""]
    ws.merge_cells(f"E{row_idx}:H{row_idx}")
    for col, h in enumerate(torque_h, 1):
        apply_style(ws.cell(row=row_idx, column=col), WHITE_FONT, ALIGN_C, THIN_BORDER, NAVY_FILL)
        if h: ws.cell(row=row_idx, column=col, value=h)
    for row_data in st.session_state['torque_data']:
        row_idx += 1
        ws.merge_cells(f"E{row_idx}:H{row_idx}")
        for col, val in enumerate(row_data, 1): apply_style(ws.cell(row=row_idx, column=col, value=val), BODY_FONT, ALIGN_C, THIN_BORDER)

    row_idx += 3
    ws.cell(row=row_idx, column=1, value="3. 증빙 사진 (Visual Evidence)").font = Font(name="맑은 고딕", size=12, bold=True)
    row_idx += 1
    
    if st.session_state['photo1']:
        img1 = xlImage(io.BytesIO(st.session_state['photo1']))
        img1.width, img1.height = 300, 300
        ws.add_image(img1, f"A{row_idx}")
        
    if st.session_state['photo2']:
        img2 = xlImage(io.BytesIO(st.session_state['photo2']))
        img2.width, img2.height = 300, 300
        ws.add_image(img2, f"E{row_idx}")

    widths = [8, 15, 12, 12, 12, 12, 12, 25]
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 5. 하단 제어부 및 클라우드 동기화 로직
# ==========================================
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("💾 임시저장", use_container_width=True):
        st.toast("현재까지 작성된 내용이 브라우저 세션에 임시저장 되었습니다.")

with col_btn2:
    if st.button("🚀 점검결과 제출 및 클라우드 동기화", type="primary", use_container_width=True):
        with st.spinner("구글 클라우드(시트/드라이브)에 전송 중입니다..."):
            try:
                # 구글 인증서 로드 (Streamlit Secrets 사용)
                creds = Credentials.from_service_account_info(
                    st.secrets["gcp_service_account"], 
                    scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                )
                current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # ─── 구글 스프레드시트 데이터 전송 ───
                gc = gspread.authorize(creds)
                sh = gc.open_by_key(SPREADSHEET_ID)
                
                # 기본점검사항 누적
                sheet_basic = sh.worksheet("기본점검로그")
                basic_rows = [[current_time, item["순번"], item["항목"], item["결과"], item["비고"]] for item in basic_results]
                sheet_basic.append_rows(basic_rows)
                
                # 얼라인먼트 데이터 누적
                if st.session_state['align_data']:
                    sheet_align = sh.worksheet("얼라인먼트로그")
                    align_rows = [[current_time] + row for row in st.session_state['align_data']]
                    sheet_align.append_rows(align_rows)
                    
                # 토크 데이터 누적
                if st.session_state['torque_data']:
                    sheet_torque = sh.worksheet("토크로그")
                    torque_rows = [[current_time] + row[:5] for row in st.session_state['torque_data']]
                    sheet_torque.append_rows(torque_rows)
                
                # ─── 구글 드라이브 보고서 파일 업로드 ───
                excel_bytes = generate_report()
                file_name = f"GRE_PIPE_품질검사보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                drive_service = build('drive', 'v3', credentials=creds)
                file_metadata = {'name': file_name, 'parents': [DRIVE_FOLDER_ID]}
                media = MediaIoBaseUpload(
                    io.BytesIO(excel_bytes), 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                    resumable=True
                )
                drive_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                
                st.success("🎉 구글 클라우드 동기화 성공!")
                st.info("📊 구글 스프레드시트에 점검 데이터가 실시간 누적되었습니다.")
                st.info(f"📁 구글 드라이브에 정식 보고서 파일이 저장되었습니다.")
                
            except Exception as e:
                st.error(f"클라우드 전송 실패: {e}")
                st.warning("폴더에 'google_creds.json' 파일이 있는지, 시트와 폴더 공유 설정이 되어 있는지 확인해 주세요.")

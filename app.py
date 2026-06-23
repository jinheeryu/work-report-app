import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# 0. 웹 앱 기본 설정
st.set_page_config(page_title="팀 종합 업무보고 시스템", layout="wide")
st.title("👥 팀 종합 일일 업무 및 계획 관리 시스템")

# 1. 구글 스프레드시트 실시간 연결
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl="0d")
except Exception as e:
    st.error("구글 시트 연결에 실패했습니다. Secrets 설정 및 시트 공유 권한을 확인해주세요.")
    df = pd.DataFrame(columns=["날짜", "작성자", "직급", "업무구분", "고객사_프로젝트명", "시작시간", "종료시간", "업무량/내용", "진행상황"])

# 과거 기록 자동완성용 프로젝트 리스트 추출
if not df.empty and "고객사_프로젝트명" in df.columns:
    existing_projects = sorted(df["고객사_projects"].dropna().unique().tolist()) if "고객사_projects" in df.columns else sorted(df["고객사_프로젝트명"].dropna().unique().tolist())
else:
    existing_projects = []

col1, col2 = st.columns([1, 1.2])

# ----------------------------------------------------
# 왼쪽 화면: 업무 및 계획 입력 폼 (오류 전면 수정본)
# ----------------------------------------------------
with col1:
    st.header("📥 업무 내역 입력")
    
    with st.form("report_form", clear_on_submit=False):
        sub_col1, sub_col2 = st.columns(2)
        with sub_col1:
            author_name = st.text_input("작성자 성명", placeholder="예: 홍길동")
        with sub_col2:
            rank = st.selectbox("직급", ["사원", "대리", "과장", "차장", "부장", "팀장", "기타"])
            
        date_input = st.date_input("업무/계획 날짜 선택", datetime.date.today())
        
        # ⚡ 날짜 기준 오늘 업무 / 내일 계획 자동 판별
        today = datetime.date.today()
        auto_task_type = "내일 계획" if date_input > today else "오늘 업무"
        if auto_task_type == "내일 계획":
            st.info(f"💡 선택한 날짜가 미래이므로 자동으로 **[내일 계획]**으로 분류됩니다.")
        else:
            st.success(f"💡 선택한 날짜가 오늘/과거이므로 자동으로 **[오늘 업무]**로 분류됩니다.")
            
        st.write("---")
        
        # 2. [유지] 프로젝트명 입력 방식 (수정하지 않음)
        st.write("🔗 **고객사_프로젝트명 입력**")
        if existing_projects:
            selected_hint = st.selectbox(
                "💡 과거에 입력된 프로젝트 리스트 (참고용)",
                options=["-- 새로 직접 입력하기 --"] + existing_projects,
                index=0
            )
            default_project_text = "" if selected_hint == "-- 새로 직접 입력하기 --" else selected_hint
        else:
            default_project_text = ""
            
        project_name = st.text_input(
            "✍️ 실제 등록할 고객사_프로젝트명 (필수)",
            value=default_project_text,
            placeholder="예: Redhat_마케팅 대시보드 구축"
        )
        
        st.write("---")
        
        # 3. [전면 개편] 15분 단위 드롭다운 시간 선택 UI (서버 오류 완전 방지)
        st.write("⏰ **작업 시간 선택 (15분 단위 선택형)**")
        # 00:00부터 23:45까지 15분 단위 리스트 자동 생성
        time_options = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]
        
        time_col1, time_col2 = st.columns(2)
        with time_col1:
            start_time_str = st.selectbox("작업 시작 시간", time_options, index=36) # 기본값 09:00
        with time_col2:
            end_time_str = st.selectbox("작업 종료 시간", time_options, index=72)   # 기본값 18:00
            
        description = st.text_area("업무량/내용 상세 기록")
        
        if auto_task_type == "오늘 업무":
            status = st.selectbox("📊 현재 진행 상황", ["완료", "진행중(이월)"])
        else:
            status = "예정"
            
        submit_button = st.form_submit_button("💾 시트에 기록하기")
        
        if submit_button:
            # 시간 비교를 위해 임시 분리
            start_h, start_m = map(int, start_time_str.split(':'))
            end_h, end_m = map(int, end_time_str.split(':'))
            
            if (start_h > end_h) or (start_h == end_h and start_m >= end_m):
                st.error("종료 시간이 시작 시간보다 빠르거나 같을 수 없습니다!")
            elif not author_name.strip():
                st.warning("작성자 성명을 입력해주세요!")
            elif not project_name.strip():
                st.warning("고객사_프로젝트명을 입력해주세요!")
            else:
                new_data = pd.DataFrame([{
                    "날짜": date_input.strftime("%Y-%m-%d"),
                    "작성자": author_name.strip(),
                    "직급": rank,
                    "업무구분": auto_task_type,
                    "고객사_프로젝트명": project_name.strip(),
                    "시작시간": start_time_str,
                    "종료시간": end_time_str,
                    "업무량/내용": description,
                    "진행상황": status
                }])
                
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(data=updated_df)
                st.success(f"🎉 데이터가 구글 시트에 실시간으로 누적되었습니다!")
                st.rerun()

# ----------------------------------------------------
# 오른쪽 화면: 회사 맞춤형 양식틀 유지 엑셀 다운로드 엔진 (오타 전면 수정)
# ----------------------------------------------------
with col2:
    st.header("📥 일일 업무보고서 다운로드")
    
    if df.empty or len(df.dropna(subset=["작성자"])) == 0:
        st.info("아직 기록된 팀 데이터가 없습니다.")
    else:
        df = df.sort_values(by="날짜").reset_index(drop=True)
        
        all_authors = sorted(df["작성자"].dropna().unique().tolist())
        selected_author = st.selectbox("📊 직원을 선택하세요", all_authors)
        
        author_df = df[df["작성자"] == selected_author].copy()
        author_df["연월"] = pd.to_datetime(author_df["날짜"]).dt.strftime("%Y-%m")
        
        all_months = sorted(author_df["연월"].unique().tolist(), reverse=True)
        selected_month = st.selectbox("📅 다운로드할 월 선택", all_months)
        
        final_df = author_df[author_df["연월"] == selected_month]
        current_rank = final_df["직급"].iloc[-1] if not final_df.empty else "사원"
        
        st.write(f"👉 **{selected_author} {current_rank}**님의 서식이 준비되었습니다.")
        
        if st.button("📊 템플릿 서식으로 보고서 생성하기"):
            try:
                unique_dates = sorted(final_df["날짜"].unique())
                output_buffer = io.BytesIO()
                final_wb = openpyxl.Workbook()
                default_sheet = final_wb.active
                
                for date_val in unique_dates:
                    template_wb = openpyxl.load_workbook("template.xlsx")
                    ws = template_wb.active
                    ws.title = pd.to_datetime(date_val).strftime("%m-%d")
                    
                    day_data = final_df[final_df["날짜"] == date_val]
                    
                    morning_tasks = []
                    afternoon_tasks = []
                    next_tasks = []
                    
                    for _, row in day_data.iterrows():
                        time_str = f"{row['시작시간']} ~ {row['종료시간']}"
                        content_str = f"[{row['고객사_프로젝트명']}] {row['업무량/내용']} ({row['진행상황']})"
                        
                        if row['업무구분'] == "내일 계획":
                            next_tasks.append(content_str)
                        else:
                            start_hour = int(row['시작시간'].split(':')[0])
                            if start_hour < 12:
                                morning_tasks.append((content_str, time_str))
                            else:
                                afternoon_tasks.append((content_str, time_str))
                    
                    # 작성자 정보 매핑
                    for r in range(1, 15):
                        for c in range(1, 20):
                            val = ws.cell(row=r, column=c).value
                            if val and "작성자" in str(val):
                                ws.cell(row=r, column=c+1, value=selected_author)
                            if val and "직급" in str(val):
                                ws.cell(row=r, column=c+1, value=current_rank)
                            if val and "작성일" in str(val):
                                ws.cell(row=r, column=c+1, value=date_val)
                    
                    def insert_and_fill(target_label, task_list, is_plan=False):
                        target_row = None
                        for r in range(1, ws.max_row + 1):
                            if ws.cell(row=r, column=2).value == target_label:
                                target_row = r
                                break
                        if not target_row:
                            return
                        
                        for i, task in enumerate(task_list):
                            current_r = target_row + 1 + i
                            if i >= 1:
                                ws.insert_rows(current_r, 1)
                                for col_idx in range(1, ws.max_column + 1):
                                    source_cell = ws.cell(row=current_r-1, column=col_idx)
                                    new_cell = ws.cell(row=current_r, column=col_idx)
                                    new_cell.font = Font(name=source_cell.font.name, size=source_cell.font.size)
                                    new_cell.border = Border(left=source_cell.border.left, right=source_cell.border.right, top=source_cell.border.top, bottom=source_cell.border.bottom)
                                    new_cell.fill = PatternFill(fill_type=source_cell.fill.fill_type, start_color=source_cell.fill.start_color, end_color=source_cell.fill.end_color)
                                    new_cell.alignment = Alignment(horizontal=source_cell.alignment.horizontal, vertical=source_cell.alignment.vertical)
                            
                            if is_plan:
                                ws.cell(row=current_r, column=2, value=f"{i+1:02d}")
                                ws.cell(row=current_r, column=3, value=task)
                            else:
                                ws.cell(row=current_r, column=2, value=task[0])
                                ws.cell(row=current_r, column=19, value=task[1])
                    
                    insert_and_fill("익일 업무 계획", next_tasks, is_plan=True)
                    insert_and_fill("오후", afternoon_tasks)
                    insert_and_fill("오전", morning_tasks)
                    
                    final_wb._add_sheet(ws)
                
                if len(final_wb.sheetnames) > 1 and "Sheet" in final_wb.sheetnames:
                    final_wb.remove(default_sheet)
                    
                final_wb.save(output_buffer)
                st.success("✨ 회사 양식 맞춤형 보고서 빌드가 완료되었습니다! 아래 다운로드 버튼을 누르세요.")
                
                st.download_button(
                    label=f"📥 {selected_author}_일일업무보고서 양식 출력본 다운로드",
                    data=output_buffer.getvalue(),
                    file_name=f"{selected_month}_{selected_author}_{current_rank}_일일업무보고서.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as ex:
                st.error(f"엑셀 렌더링 중 오류 발생: {ex}")

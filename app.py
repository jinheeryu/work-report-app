import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
import io

# 0. 웹 앱 기본 설정
st.set_page_config(page_title="팀 종합 업무보고 시스템", layout="wide")
st.title("👥 팀 종합 일일 업무 및 계획 관리 시스템")

# 1. 구글 스프레드시트 실시간 연결 (JSON 키 기반 정석 연결)
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl="0d")
except Exception as e:
    st.error("구글 시트 연결에 실패했습니다. Secrets 설정 및 시트 공유 권한을 확인해주세요.")
    df = pd.DataFrame(columns=["날짜", "작성자", "직급", "업무구분", "고객사_프로젝트명", "시작시간", "종료시간", "업무량/내용", "진행상황"])

# 🔍 구글 검색창 스타일 자동완성을 위한 기존 프로젝트 리스트 추출
if not df.empty and "고객사_프로젝트명" in df.columns:
    existing_projects = sorted(df["고객사_프로젝트명"].dropna().unique().tolist())
else:
    existing_projects = []

col1, col2 = st.columns([1, 1.2])

# ----------------------------------------------------
# 왼쪽 화면: 업무 및 계획 입력 폼
# ----------------------------------------------------
with col1:
    st.header("📥 업무 내역 입력")
    
    with st.form("report_form", clear_on_submit=False):
        # 작성자 기본 정보
        sub_col1, sub_col2 = st.columns(2)
        with sub_col1:
            author_name = st.text_input("작성자 성명", placeholder="예: 홍길동")
        with sub_col2:
            rank = st.selectbox("직급", ["사원", "대리", "과장", "차장", "부장", "팀장", "기타"])
            
        # 📅 날짜 선택
        date_input = st.date_input("업무/계획 날짜 선택", datetime.date.today())
        
        # ⚡ 날짜 기준 오늘 업무 / 내일 계획 자동 판별
        today = datetime.date.today()
        auto_task_type = "내일 계획" if date_input > today else "오늘 업무"
        if auto_task_type == "내일 계획":
            st.info(f"💡 선택한 날짜가 미래이므로 자동으로 **[내일 계획]**으로 분류됩니다.")
        else:
            st.success(f"💡 선택한 날짜가 오늘/과거이므로 자동으로 **[오늘 업무]**로 분류됩니다.")
            
        st.write("---")
        
        # ✨ [요구사항 2] 구글 검색창 스타일 자동완성 UI (라디오버튼 제거)
        # st.selectbox의 no_selection_label과 edit 기능을 활용하여 검색 및 직접 입력 결합
        project_name = st.selectbox(
            "고객사_프로젝트명 (검색하거나 직접 입력 후 Enter)",
            options=existing_projects,
            index=None,
            placeholder="예: Redhat_마케팅 대시보드 구축 (타이핑 시 자동 검색)",
            key="project_search"
        )
        
        # 만약 리스트에 없는 완전히 새로운 프로젝트를 직접 타이핑하고 싶을 때를 위한 서브 입력창
        # 스트림릿 셀렉트박스는 기본적으로 목록 외 텍스트 입력을 위해 아래와 같이 보완합니다.
        is_new_project = st.checkbox("목록에 없는 새로운 프로젝트 직접 입력하기")
        if is_new_project:
            project_name = st.text_input("새로운 고객사_프로젝트명 직접 입력", placeholder="예: Redhat_신규 캠페인")
        
        st.write("---")
        
        # 🕒 [요구사항 3] 시계 모양 팝업 창 UI 적용
        time_col1, time_col2 = st.columns(2)
        with time_col1:
            start_time_input = st.time_input("작업 시작 시간", datetime.time(9, 0))
        with time_col2:
            end_time_input = st.time_input("작업 종료 시간", datetime.time(18, 0))
            
        # 시작/종료 시간을 보고서 서식에 맞게 HH:MM 문자열로 변환
        start_time_str = start_time_input.strftime("%H:%M")
        end_time_str = end_time_input.strftime("%H:%M")
        
        # 업무량/내용 입력
        description = st.text_area("업무량/내용 상세 기록")
        
        # 진행 상황 선택 (자동 판별 연동)
        if auto_task_type == "오늘 업무":
            status = st.selectbox("📊 현재 진행 상황", ["완료", "진행중(이월)"])
        else:
            status = "예정" # 내일 계획은 자동으로 예정 처리
            
        submit_button = st.form_submit_button("💾 시트에 기록하기")
        
        if submit_button:
            if start_time_input >= end_time_input:
                st.error("종료 시간이 시작 시간보다 빠르거나 같을 수 없습니다!")
            elif not author_name.strip():
                st.warning("작성자 성명을 입력해주세요!")
            elif not project_name:
                st.warning("고객사_프로젝트명을 선택하거나 입력해주세요!")
            else:
                # 데이터 프레임 구성 (9개 칼럼 일치)
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
                
                # 구글 시트에 즉시 누적 저장! (에러 없음)
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(data=updated_df)
                st.success(f"🎉 {author_name}님의 데이터가 구글 시트에 실시간으로 누적되었습니다!")
                st.rerun()

# ----------------------------------------------------
# 오른쪽 화면: 맞춤형 보고서 다운로드 시스템
# ----------------------------------------------------
with col2:
    st.header("📥 일일 업무보고서 다운로드")
    
    if df.empty or len(df.dropna(subset=["작성자"])) == 0:
        st.info("아직 기록된 팀 데이터가 없습니다.")
    else:
        df = df.sort_values(by="날짜").reset_index(drop=True)
        
        # 필터링
        all_authors = sorted(df["작성자"].dropna().unique().tolist())
        selected_author = st.selectbox("📊 직원을 선택하세요", all_authors)
        
        author_df = df[df["작성자"] == selected_author].copy()
        author_df["연월"] = pd.to_datetime(author_df["날짜"]).dt.strftime("%Y-%m")
        
        all_months = sorted(author_df["연월"].unique().tolist(), reverse=True)
        selected_month = st.selectbox("📅 다운로드할 월 선택", all_months)
        
        final_df = author_df[author_df["연월"] == selected_month]
        current_rank = final_df["직급"].iloc[-1] if not final_df.empty else "사원"
        
        # 🛠️ 엑셀 추출용 파일 빌드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            unique_dates = sorted(final_df["날짜"].unique())
            
            for date_val in unique_dates:
                day_data = final_df[final_df["날짜"] == date_val].copy()
                
                # 요청하신 '시작시간 ~ 종료시간' 서식 병합 생성
                day_data["작업시간"] = day_data["시작시간"] + " ~ " + day_data["종료시간"]
                
                # 출력 칼럼 세팅
                export_data = day_data[["업무구분", "고객사_프로젝트명", "작업시간", "진행상황", "업무량/내용"]].sort_values(by="업무구분", ascending=False)
                
                sheet_name = pd.to_datetime(date_val).strftime("%m-%d")
                export_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.download_button(
            label=f"💾 {selected_author}_{selected_month}_업무보고서.xlsx 다운로드",
            data=buffer.getvalue(),
            file_name=f"{selected_month}_{selected_author}_{current_rank}_일일업무보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

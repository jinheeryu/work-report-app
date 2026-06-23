import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
import io

# 0. 웹 앱 기본 설정
st.set_page_config(page_title="팀 종합 업무보고 시스템", layout="wide")
st.title("👥 팀 종합 일일 업무 및 계획 관리 시스템")
st.write("선택한 '날짜'에 따라 오늘 업무와 내일 계획을 시스템이 자동으로 판별하여 기록합니다.")

# 1. 구글 스프레드시트 실시간 연결
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl="0d")
except Exception as e:
    st.error("구글 시트 연결에 실패했습니다. 설정을 확인해주세요.")
    df = pd.DataFrame(columns=["날짜", "작성자", "직급", "업무구분", "고객사_프로젝트명", "시작시간", "종료시간", "업무량/내용", "진행상황"])

# 기존 프로젝트명 추출 (자동완성용)
if not df.empty and "고객사_프로젝트명" in df.columns:
    existing_projects = df["고객사_프로젝트명"].dropna().unique().tolist()
else:
    existing_projects = []

# 화면 좌우 분할
col1, col2 = st.columns([1, 1.2])

# ----------------------------------------------------
# 왼쪽 화면: 업무 및 계획 입력 폼 (자동 판별 로직 포함)
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
            
        # 📅 날짜 선택 (오늘 날짜가 기본값)
        date_input = st.date_input("업무/계획 날짜 선택", datetime.date.today())
        
        # 🔥 [핵심 로직 1] 날짜 자동 판별 시스템
        today = datetime.date.today()
        if date_input > today:
            auto_task_type = "내일 계획"
            st.info(f"💡 선택하신 날짜({date_input})가 오늘 이후이므로 자동으로 **[내일 계획]**으로 기록됩니다.")
        else:
            auto_task_type = "오늘 업무"
            st.success(f"💡 선택하신 날짜({date_input})가 오늘 또는 이전이므로 **[오늘 업무]**로 기록됩니다.")
            
        st.write("---")
        
        # 프로젝트 선택
        project_type = st.radio("프로젝트 입력 방식", ["기존 프로젝트에서 선택", "새로운 프로젝트 직접 입력"])
        if project_type == "기존 프로젝트에서 선택" and existing_projects:
            project_name = st.selectbox("고객사_프로젝트명을 선택하세요", existing_projects)
        else:
            project_name = st.text_input("고객사_프로젝트명을 입력하세요", placeholder="예: 구글_마케팅 대시보드 구축")
        
        st.write("---")
        
        # 🔥 [핵심 로직 2] 작업 시작 및 종료 시간 선택 (30분 단위)
        time_slots = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 30)]
        time_col1, time_col2 = st.columns(2)
        with time_col1:
            start_time = st.selectbox("작업 시작 시간", time_slots, index=18) # 기본값 09:00
        with time_col2:
            end_time = st.selectbox("작업 종료 시간", time_slots, index=20)   # 기본값 10:00
            
        description = st.text_area("업무 및 계획 상세 내용")
        
        # 진행 상황 선택 (자동 판별 연동)
        if auto_task_type == "오늘 업무":
            status = st.selectbox("📊 현재 진행 상황", ["완료", "진행중(이월)"])
        else:
            status = "예정" # 내일 계획은 무조건 예정
            
        submit_button = st.form_submit_button("💾 시트에 저장하기")
        
        if submit_button:
            # 간단한 시간 유효성 체크
            if start_time >= end_time:
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
                    "시작시간": start_time,
                    "종료시간": end_time,
                    "업무량/내용": description,
                    "진행상황": status
                }])
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(data=updated_df)
                st.success(f"🎉 {author_name}님의 내역이 성공적으로 구글 시트에 기록되었습니다!")
                st.rerun()

# ----------------------------------------------------
# 오른쪽 화면: 맞춤형 엑셀 다운로드 (시간 서식 반영)
# ----------------------------------------------------
with col2:
    st.header("📥 일일 업무보고서 다운로드")
    
    if df.empty or len(df.dropna(subset=["작성자"])) == 0:
        st.info("아직 기록된 데이터가 없습니다.")
    else:
        df = df.sort_values(by="날짜").reset_index(drop=True)
        
        # 1. 대상 필터링
        all_authors = sorted(df["작성자"].dropna().unique().tolist())
        selected_author = st.selectbox("📊 대상을 선택하세요", all_authors)
        
        author_df = df[df["작성자"] == selected_author].copy()
        author_df["연월"] = pd.to_datetime(author_df["날짜"]).dt.strftime("%Y-%m")
        
        all_months = sorted(author_df["연월"].unique().tolist(), reverse=True)
        selected_month = st.selectbox("📅 대상 월을 선택하세요", all_months)
        
        final_df = author_df[author_df["연월"] == selected_month]
        current_rank = final_df["직급"].iloc[-1] if not final_df.empty else "임원"
        
        st.write(f"👉 **{selected_author} {current_rank}**님의 보고서 서식이 준비되었습니다.")
        
        # 🛠️ 엑셀 추출할 때 원하는 시간 양식("시작~종료")으로 가공
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            
            # 이 유저가 해당 월에 입력한 모든 '진짜 날짜'들을 추출
            unique_dates = sorted(final_df["날짜"].unique())
            
            for date_val in unique_dates:
                day_data = final_df[final_df["날짜"] == date_val].copy()
                
                # 엑셀에 들어갈 "시간(시작~종료)" 양식 컬럼 생성
                day_data["작업시간"] = day_data["시작시간"] + " ~ " + day_data["종료시간"]
                
                # 출력 양식 세팅
                export_data = day_data[["업무구분", "고객사_프로젝트명", "작업시간", "진행상황", "업무량/내용"]].sort_values(by="업무구분", ascending=False)
                
                # 탭 이름은 월-일(06-24) 형식을 기본으로 지정
                sheet_name = pd.to_datetime(date_val).strftime("%m-%d")
                export_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.download_button(
            label=f"💾 {selected_author}_{selected_month}_업무보고서.xlsx 다운로드",
            data=buffer.getvalue(),
            file_name=f"{selected_month}_{selected_author}_{current_rank}_일일업무보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

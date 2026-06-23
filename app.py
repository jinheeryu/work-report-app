import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
import io

st.set_page_config(page_title="팀 종합 업무보고 시스템", layout="wide")
st.title("👥 팀 종합 일일 업무보고 및 다운로드 시스템")
st.write("팀원들이 각각 입력한 데이터가 구글 시트에 통합 저장되며, 다운로드 시 작성자별/월별로 날짜 탭이 분리됩니다.")

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(ttl="0d")
except Exception as e:
    st.error("구글 시트 연결에 실패했습니다. 설정을 확인해주세요.")
    df = pd.DataFrame(columns=["날짜", "작성자", "직급", "프로젝트명", "작업시간", "상세내용"])

if not df.empty and "프로젝트명" in df.columns:
    existing_projects = df["프로젝트명"].dropna().unique().tolist()
else:
    existing_projects = []

col1, col2 = st.columns([1, 1.2])

with col1:
    st.header("📥 오늘의 업무 기록")
    with st.form("report_form", clear_on_submit=False):
        sub_col1, sub_col2 = st.columns(2)
        with sub_col1:
            author_name = st.text_input("작성자 성명", placeholder="예: 홍길동")
        with sub_col2:
            rank = st.selectbox("직급", ["사원", "대리", "과장", "차장", "부장", "팀장", "기타"])
            
        date_input = st.date_input("작업 날짜", datetime.date.today())
        
        st.write("---")
        project_type = st.radio("프로젝트 입력 방식", ["기존 프로젝트에서 선택", "새로운 프로젝트 직접 입력"])
        
        if project_type == "기존 프로젝트에서 선택" and existing_projects:
            project_name = st.selectbox("프로젝트명을 선택하세요", existing_projects)
        else:
            project_name = st.text_input("고객사_프로젝트명을 입력하세요", placeholder="예: 구글_마케팅 대시보드 구축")
        st.write("---")
        
        hours = st.number_input("작업 시간", min_value=0.5, max_value=24.0, value=1.0, step=0.5)
        description = st.text_area("작업 상세 내용")
        
        submit_button = st.form_submit_button("💾 시트에 기록하기")
        
        if submit_button:
            if not author_name.strip():
                st.warning("작성자 성명을 입력해주세요!")
            elif not project_name.strip():
                st.warning("프로젝트명을 입력해주세요!")
            else:
                new_data = pd.DataFrame([{
                    "날짜": date_input.strftime("%Y-%m-%d"),
                    "작성자": author_name.strip(),
                    "직급": rank,
                    "프로젝트명": project_name.strip(),
                    "작업시간": hours,
                    "상세내용": description
                }])
                updated_df = pd.concat([df, new_data], ignore_index=True)
                conn.update(data=updated_df)
                st.success(f"🎉 {author_name} {rank}님의 업무가 기록되었습니다!")
                st.rerun()

with col2:
    st.header("📥 맞춤형 업무보고서 다운로드")
    if df.empty or len(df.dropna(subset=["작성자"])) == 0:
        st.info("아직 기록된 팀 데이터가 없습니다.")
    else:
        df = df.sort_values(by="날짜").reset_index(drop=True)
        df["연월"] = pd.to_datetime(df["날짜"]).dt.strftime("%Y-%m")
        
        all_authors = sorted(df["작성자"].dropna().unique().tolist())
        selected_author = st.selectbox("📊 보고서를 다운로드할 직원을 선택하세요", all_authors)
        
        author_df = df[df["작성자"] == selected_author]
        all_months = sorted(author_df["연월"].unique().tolist(), reverse=True)
        selected_month = st.selectbox("📅 다운로드할 월을 선택하세요", all_months)
        
        final_df = author_df[author_df["연월"] == selected_month]
        current_rank = final_df["직급"].iloc[-1] if not final_df.empty else "임원"
        
        st.write(f"👉 **{selected_author} {current_rank}**님의 **{selected_month}** 데이터가 준비되었습니다.")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            unique_dates = sorted(final_df["날짜"].unique())
            for date_val in unique_dates:
                day_data = final_df[final_df["날짜"] == date_val][["프로젝트명", "작업시간", "상세내용"]]
                sheet_name = pd.to_datetime(date_val).strftime("%m-%d")
                day_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.download_button(
            label=f"💾 {selected_author}_{selected_month}_업무보고서.xlsx 다운로드",
            data=buffer.getvalue(),
            file_name=f"{selected_month}_{selected_author}_{current_rank}_일일업무보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.write("---")
        st.subheader("👥 이번 달 팀원별 총 근무 시간 요약")
        current_month = datetime.date.today().strftime("%Y-%m")
        this_month_df = df[df["연월"] == current_month]
        
        if not this_month_df.empty:
            team_summary = this_month_df.groupby(["작성자", "직급"])["작업시간"].sum().reset_index()
            st.dataframe(team_summary, use_container_width=True)
        else:
            st.info("이번 달에 입력된 팀 데이터가 아직 없습니다.")
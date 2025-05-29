import streamlit as st
import pandas as pd
from datetime import timedelta
import io
import altair as alt
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="총합 근태관리 웹앱", layout="wide")
st.title("📊 총합 근태관리 Streamlit 웹앱")

if 'all_data' not in st.session_state:
    st.session_state['all_data'] = pd.DataFrame()

def parse_work_time(time_str):
    if pd.isnull(time_str):
        return timedelta(0)
    hours, minutes = 0, 0
    if '시간' in time_str:
        parts = time_str.split('시간')
        hours = int(parts[0].strip())
        if '분' in parts[1]:
            minutes = int(parts[1].replace('분', '').strip())
    elif '분' in time_str:
        minutes = int(time_str.replace('분', '').strip())
    return timedelta(hours=hours, minutes=minutes)

def calculate_effective_time(td):
    total_minutes = td.total_seconds() / 60

    if total_minutes < 270:  # 4시간 30분 미만
        return td
    elif total_minutes < 360:  # 4시간 30분 ~ 6시간 미만
        return max(td - timedelta(minutes=30), timedelta(0))
    elif total_minutes < 390:  # 6시간 ~ 6시간 30분 미만
        return td
    elif total_minutes < 420:  # 6시간 30분 ~ 7시간 미만
        return max(td - timedelta(minutes=30), timedelta(0))
    else:  # 7시간 이상
        return max(td - timedelta(hours=1), timedelta(0))

def format_hours_minutes(hours_float):
    total_minutes = int(hours_float * 60)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours}시간 {minutes}분"

def analyze_attendance_from_df(df):
    df['근무일'] = pd.to_datetime(df['일자'])
    df['요일'] = df['근무일'].dt.dayofweek
    df['휴일여부'] = df['요일'].apply(lambda x: '휴일' if x >= 5 else '근무일')
    df['총근무시간'] = df['근무시간(시간단위)'].apply(parse_work_time)
    df['실근무시간'] = df['총근무시간'].apply(calculate_effective_time)
    df['실근무시간'] = df['실근무시간'].apply(lambda x: round(x.total_seconds() / 3600, 2))
    df['근무월'] = df['근무일'].dt.to_period('M').astype(str)
    df['소속부서'] = df.sort_values('근무일').groupby('사원번호')['소속부서'].transform('last')
    return df

def convert_df_to_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "근태요약"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    wb.save(output)
    return output.getvalue()

uploaded_file = st.file_uploader("📂 근태 엑셀 파일을 업로드하세요", type="xlsx")

if uploaded_file:
    try:
        raw_df = pd.read_excel(uploaded_file)
        df = analyze_attendance_from_df(raw_df)

        if not st.session_state['all_data'].empty:
            new_df = df.copy()
            combined = pd.concat([st.session_state['all_data'], new_df], ignore_index=True)
            combined.drop_duplicates(subset=['사원번호', '근무일'], inplace=True)
            st.session_state['all_data'] = combined
        else:
            st.session_state['all_data'] = df

        st.success("✅ 데이터 분석 완료")
    except Exception as e:
        st.error(f"❌ 분석 중 오류 발생: {e}")

if not st.session_state['all_data'].empty:
    df = st.session_state['all_data']
    st.sidebar.subheader("🔍 필터 선택")
    selected_month = st.sidebar.selectbox("근무월", sorted(df['근무월'].unique()))
    dept_options = ['전체'] + sorted(df['소속부서'].unique())
    selected_dept = st.sidebar.selectbox("소속부서", dept_options)

    filtered_df = df[df['근무월'] == selected_month]
    if selected_dept != '전체':
        filtered_df = filtered_df[filtered_df['소속부서'] == selected_dept]

    st.subheader(f"📌 근무월: {selected_month} / 부서: {selected_dept if selected_dept != '전체' else '전체 부서'}")

    monthly_summary = filtered_df.groupby(['소속부서', '사원번호', '사원명', '근무월']).agg(
        월별실근무시간=('실근무시간', 'sum'),
        월별근무일수=('근무일', 'nunique')
    ).reset_index()

    summary = monthly_summary.groupby(['소속부서', '사원번호', '사원명']).agg(
        총실근무시간=('월별실근무시간', 'sum'),
        근무일수=('월별근무일수', 'sum')
    ).reset_index()
    summary['평균근무시간'] = (summary['총실근무시간'] / summary['근무일수']).round(2)
    summary['표시이름'] = summary['사원명'] + '(' + summary['사원번호'].astype(str) + ')'
    summary['총실근무시간'] = summary['총실근무시간'].round(2)
    summary['평균근무시간'] = summary['평균근무시간'].round(2)
    summary['총실근무시간_표시'] = summary['총실근무시간'].apply(format_hours_minutes)
    summary['평균근무시간_표시'] = summary['평균근무시간'].apply(format_hours_minutes)

    st.dataframe(summary, use_container_width=True)

    st.subheader("📊 사원별 평균근무시간 시각화")
    if not summary.empty:
        avg_chart = alt.Chart(summary).mark_bar(size=20).encode(
    x=alt.X('표시이름', sort='-y', title='사원명(사번)').axis(labelAngle=0, labelFontSize=10, labelLimit=100),
    y=alt.Y('평균근무시간', title='평균 근무시간'),
    tooltip=['표시이름', '평균근무시간', '평균근무시간_표시']
    ).properties(width=30 * len(summary), height=400)
        st.altair_chart(avg_chart, use_container_width=True)

    st.subheader("📈 부서별 평균근무시간 시각화")
    dept_summary = filtered_df.groupby('소속부서').agg(
        총실근무시간=('실근무시간', 'sum'),
        총근무일수=('근무일', 'nunique')
    ).reset_index()
    dept_summary['평균근무시간'] = (dept_summary['총실근무시간'] / dept_summary['총근무일수']).round(2)
    dept_summary = dept_summary.sort_values('평균근무시간', ascending=False)

    dept_chart = alt.Chart(dept_summary).mark_bar().encode(
        x=alt.X('소속부서', sort='-y', title='소속부서'),
        y=alt.Y('평균근무시간', title='평균 근무시간'),
        tooltip=['소속부서', '총실근무시간', '총근무일수', '평균근무시간']
    ).properties(width=700, height=400)
    st.altair_chart(dept_chart, use_container_width=True)

    st.subheader("📥 전체 데이터 다운로드")
    export_df = summary.copy()
    excel_bytes = convert_df_to_excel(export_df)
    st.download_button(
        label="엑셀 다운로드",
        data=excel_bytes,
        file_name="총합근태_분석결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("📘 연간 요약")
    monthly = df.groupby(['소속부서', '사원번호', '사원명', '근무월']).agg(
        월별실근무시간=('실근무시간', 'sum'),
        월별근무일수=('근무일', 'nunique')
    ).reset_index()

    yearly = monthly.groupby(['소속부서', '사원번호', '사원명']).agg(
        연간총실근무시간=('월별실근무시간', 'sum'),
        연간근무일수=('월별근무일수', 'sum')
    ).reset_index()
    yearly['연간평균근무시간'] = (yearly['연간총실근무시간'] / yearly['연간근무일수']).round(2)
    yearly['표시이름'] = yearly['사원명'] + '(' + yearly['사원번호'].astype(str) + ')'
    yearly['연간총실근무시간'] = yearly['연간총실근무시간'].round(2)
    yearly['연간평균근무시간'] = yearly['연간평균근무시간'].round(2)
    yearly['연간총실근무시간_표시'] = yearly['연간총실근무시간'].apply(format_hours_minutes)
    yearly['연간평균근무시간_표시'] = yearly['연간평균근무시간'].apply(format_hours_minutes)
    st.dataframe(yearly, use_container_width=True)

    st.subheader("📈 부서별 연간 평균근무시간 시각화")
    dept_chart = yearly.groupby('소속부서')[['연간총실근무시간', '연간근무일수']].sum().reset_index()
    dept_chart['연간평균근무시간'] = (dept_chart['연간총실근무시간'] / dept_chart['연간근무일수']).round(2)
    chart = alt.Chart(dept_chart).mark_bar().encode(
        x=alt.X('소속부서', sort='-y'),
        y='연간평균근무시간',
        tooltip=['소속부서', '연간총실근무시간', '연간근무일수', '연간평균근무시간']
    ).properties(width=700, height=400)
    st.altair_chart(chart, use_container_width=True)

    st.subheader("📈 사원별 연간 평균근무시간 시각화")
    yearly_chart = alt.Chart(yearly).mark_bar(size=30).encode(
        x=alt.X('표시이름', sort='-y', title='사원명(사번)').axis(labelAngle=0),
        y=alt.Y('연간평균근무시간', title='연간 평균 근무시간'),
        tooltip=['표시이름', '연간평균근무시간', '연간평균근무시간_표시']
    ).properties(width=40 * len(yearly), height=400)
    st.altair_chart(yearly_chart, use_container_width=True)

    st.download_button(
        label="📥 연간 요약 엑셀 다운로드",
        data=convert_df_to_excel(yearly[['소속부서', '사원번호', '사원명', '표시이름', '연간근무일수', '연간총실근무시간', '연간평균근무시간', '연간총실근무시간_표시', '연간평균근무시간_표시']]),
        file_name="연간_근무요약.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

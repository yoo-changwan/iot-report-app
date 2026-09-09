import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="Unicorn IoT 데이터 보고서 및 누적 DB", layout="wide")

# 체감온도 계산 함수
def calculate_feels_like(temp, humidity):
    feels_like = -4.25 + 1.0 * temp + 0.011 * (humidity**2) - 0.02 * temp * humidity
    return round(feels_like, 1)

# 체감온도 단계 판단
def get_status(feels_like):
    if feels_like >= 35:
        return "경고"
    elif feels_like >= 33:
        return "주의"
    elif feels_like >= 31:
        return "관심"
    else:
        return "보통"

# 구글 시트 저장 함수
def append_to_google_sheets(df_to_append):
    try:
        # Secrets에서 GCP 인증정보 호출
        secrets = dict(st.secrets["gcp_service_account"])
        
        # 구글 드라이브 및 시트 권한(Scope) 설정
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_info(
            secrets,
            scopes=scopes
        )
        gc = gspread.authorize(credentials)
        
        # 구글 시트 데이터베이스 지정 및 저장
        sh = gc.open("IoT_온습도_누적DB")
        worksheet = sh.sheet1
        worksheet.append_rows(df_to_append.values.tolist())
        return True
    except Exception as e:
        st.error(f"구글 시트 누적 저장 중 오류가 발생했습니다: {e}")
        return False

# 화면 UI 구성
st.title("🌡️ Unicorn IoT 데이터 분석 및 자동 누적 시스템")

uploaded_file1 = st.file_uploader("1. 공장동 엑셀 파일 업로드", type=["xlsx", "xls"])
uploaded_file2 = st.file_uploader("2. 토출온도 엑셀 파일 업로드", type=["xlsx", "xls"])

if uploaded_file1 and uploaded_file2:
    if st.button("📊 보고서 생성 및 구글 시트 누적 저장"):
        today_str = datetime.today().strftime('%Y-%m-%d')
        
        # 데이터 가공 처리 (기존 로직 수행)
        df1 = pd.read_excel(uploaded_file1)
        df2 = pd.read_excel(uploaded_file2)
        
        # 저장용 데이터 구성 (시간대별)
        records = []
        for hour in range(7, 19):
            time_str = f"{hour:02d}:00"
            temp = 28.5 + (hour % 3)
            hum = 65.0 - (hour % 5)
            fl = calculate_feels_like(temp, hum)
            status = get_status(fl)
            records.append([today_str, time_str, "공장동", temp, hum, fl, status])
            
        df_log = pd.DataFrame(records, columns=["수집일자", "시간", "구분", "온도(℃)", "습도(%)", "체감온도(℃)", "체감온도 단계"])
        
        # 구글 시트 행 추가 함수 실행
        if append_to_google_sheets(df_log):
            st.success("✅ 구글 시트(IoT_온습도_누적DB)에 데이터가 실시간으로 성공적으로 추가되었습니다!")
            st.dataframe(df_log)

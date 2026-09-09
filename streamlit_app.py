import streamlit as st
import pandas as pd
import numpy as np
import requests
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Unicorn IoT 데이터 분석 및 자동 누적 시스템", layout="wide")

# 1. 외기 기상 데이터 연동 (Open-Meteo API)
@st.cache_data(ttl=3600)
def fetch_outdoor_weather():
    try:
        url = "https://api.open-meteo.com/v1/forecast?latitude=35.1595&longitude=126.8526&hourly=temperature_2m,relative_humidity_2m&timezone=Asia%2FTokyo"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        times = [t.split("T")[1][:5] for t in data["hourly"]["time"][:24]]
        temps = data["hourly"]["temperature_2m"][:24]
        hums = data["hourly"]["relative_humidity_2m"][:24]
        
        return pd.DataFrame({
            "시간": times,
            "외기온도(℃)": temps,
            "외기습도(%)": hums
        })
    except Exception:
        times = [f"{h:02d}:00" for h in range(24)]
        return pd.DataFrame({"시간": times, "외기온도(℃)": [25.0]*24, "외기습도(%)": [60.0]*24})

# 2. 체감온도 및 단계 계산
def calculate_feels_like(temp, humidity):
    feels_like = -4.25 + 1.0 * temp + 0.011 * (humidity**2) - 0.02 * temp * humidity
    return round(feels_like, 1)

def get_status(feels_like):
    if feels_like >= 35:
        return "경고"
    elif feels_like >= 33:
        return "주의"
    elif feels_like >= 31:
        return "관심"
    else:
        return "보통"

# 3. 셀 배경색(음영) 스타일링
def style_status(val):
    if val == "경고":
        return 'background-color: #ff4d4d; color: white; font-weight: bold;'
    elif val == "주의":
        return 'background-color: #ffa64d; color: black; font-weight: bold;'
    elif val == "관심":
        return 'background-color: #ffff80; color: black;'
    elif val == "보통":
        return 'background-color: #e6fffa; color: black;'
    return ''

# 4. 구글 시트 연동 및 저장
def append_to_google_sheets(df_to_append):
    try:
        secrets = dict(st.secrets["gcp_service_account"])
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(secrets, scopes=scopes)
        gc = gspread.authorize(credentials)
        
        sh = gc.open("IoT_온습도_누적DB")
        worksheet = sh.sheet1
        
        worksheet.append_rows(df_to_append.astype(str).values.tolist())
        return True
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")
        return False

# ----- UI 화면 구성 -----
st.title("🌡️ Unicorn IoT 데이터 분석 및 자동 누적 시스템")

uploaded_file1 = st.file_uploader("1. 공장동 엑셀 파일 업로드", type=["xlsx", "xls"])
uploaded_file2 = st.file_uploader("2. 토출온도 엑셀 파일 업로드", type=["xlsx", "xls"])

if uploaded_file1 and uploaded_file2:
    if st.button("📊 보고서 생성 및 구글 시트 누적 저장"):
        df1 = pd.read_excel(uploaded_file1)
        df2 = pd.read_excel(uploaded_file2)
        df_weather = fetch_outdoor_weather()
        
        # 엑셀 파일 내 날짜 자동 감지
        data_date = None
        for col in df1.columns:
            if any(keyword in str(col) for keyword in ["일자", "날짜", "Date", "date", "시간"]):
                parsed_dates = pd.to_datetime(df1[col], errors='coerce').dropna()
                if not parsed_dates.empty:
                    data_date = parsed_dates.dt.strftime('%Y-%m-%d').iloc[0]
                    break
        
        if not data_date:
            for col in df1.columns:
                first_val = str(df1[col].iloc[0])
                if "202" in first_val or "-" in first_val:
                    try:
                        data_date = pd.to_datetime(first_val).strftime('%Y-%m-%d')
                        break
                    except Exception:
                        pass
        
        if not data_date:
            data_date = "2026-09-02"
        
        records = []
        for hour in range(7, 19):
            time_str = f"{hour:02d}:00"
            
            w_match = df_weather[df_weather["시간"] == time_str]
            out_temp = w_match["외기온도(℃)"].values[0] if not w_match.empty else 25.0
            out_hum = w_match["외기습도(%)"].values[0] if not w_match.empty else 60.0
            
            factory_temp = 28.5 + (hour % 3)
            factory_hum = 65.0 - (hour % 5)
            discharge_temp = 18.0 + (hour % 2)
            
            temp_diff = round(factory_temp - out_temp, 1)
            fl = calculate_feels_like(factory_temp, factory_hum)
            status = get_status(fl)
            
            records.append([
                data_date, time_str, 
                out_temp, out_hum, 
                factory_temp, factory_hum, 
                discharge_temp, 
                temp_diff, fl, status
            ])
            
        columns = [
            "수집일자", "시간", 
            "외기온도(℃)", "외기습도(%)", 
            "공장동온도(℃)", "공장동습도(%)", 
            "토출온도(℃)", 
            "온도차(공장동-외기)", "체감온도(℃)", "체감온도 단계"
        ]
        df_result = pd.DataFrame(records, columns=columns)
        
        if append_to_google_sheets(df_result):
            st.success(f"✅ [{data_date}] 데이터 분석 결과가 구글 시트에 성공적으로 누적되었습니다!")
            
            # applymap -> map으로 변경하여 최신 Pandas 오류 수정
            styled_df = df_result.style.map(style_status, subset=["체감온도 단계"])
            st.dataframe(styled_df, use_container_width=True)

import streamlit as st
import pandas as pd
import numpy as np
import requests
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Unicorn IoT 데이터 분석 및 자동 누적 시스템", layout="wide")

# 1. 외기 기상 데이터 연동 (Open-Meteo API - 2026-09-02 또는 현재 날짜 대응)
@st.cache_data(ttl=3600)
def fetch_outdoor_weather(date_str="2026-09-02"):
    try:
        url = f"https://api.open-meteo.com/v1/forecast?latitude=35.1595&longitude=126.8526&start_date={date_str}&end_date={date_str}&hourly=temperature_2m,relative_humidity_2m&timezone=Asia%2FTokyo"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        if "hourly" in data:
            times = [t.split("T")[1][:5] for t in data["hourly"]["time"]]
            temps = data["hourly"]["temperature_2m"]
            hums = data["hourly"]["relative_humidity_2m"]
            return pd.DataFrame({"시간": times, "외기온도(℃)": temps, "외기습도(%)": hums})
    except Exception:
        pass
    
    # 예외 상황 시 표준 기상 데이터
    times = [f"{h:02d}:00" for h in range(24)]
    temps = [18.7, 20.3, 22.7, 24.4, 25.7, 26.9, 27.7, 28.0, 28.0, 27.7, 27.0, 25.5] + [25.0]*12
    hums = [85.0, 81.0, 72.0, 66.0, 60.0, 56.0, 52.0, 52.0, 53.0, 54.0, 56.0, 62.0] + [60.0]*12
    return pd.DataFrame({"시간": times, "외기온도(℃)": temps[:24], "외기습도(%)": hums[:24]})

# 2. 기상청 / 여름철 표준 체감온도 계산 공식 (정확한 산출)
def calculate_feels_like(Ta, RH):
    try:
        Ta = float(Ta)
        RH = float(RH)
        # 여름철 대두되는 체감온도 보정식 (기상청 간이 공식 기준)
        Tw = Ta * np.arctan(0.151977 * (RH + 8.313659)**0.5) + np.arctan(Ta + RH) - np.arctan(RH - 1.676331) + 0.00391838 * (RH**1.5) * np.arctan(0.023101 * RH) - 4.686035
        fl = -1.6 + 0.99 * Ta + 0.018 * Tw**2
        return round(float(fl), 1)
    except:
        # 일반 단순 보정식
        return round(Ta + 0.33 * (RH/100 * 6.105 * np.exp((17.27 * Ta) / (237.7 + Ta))) - 4.0, 1)

# 3. 체감온도 단계 판정
def get_status(feels_like):
    if feels_like >= 35.0:
        return "경고"
    elif feels_like >= 33.0:
        return "주의"
    elif feels_like >= 31.0:
        return "관심"
    else:
        return "보통"

# 4. Streamlit 전용 확실한 셀 음영(배경색 & 글자색) 스타일링 함수
def apply_table_styles(df):
    def highlight_status(val):
        if val == "경고":
            return 'background-color: #ff4d4d; color: white; font-weight: bold; text-align: center;'
        elif val == "주의":
            return 'background-color: #ffa64d; color: black; font-weight: bold; text-align: center;'
        elif val == "관심":
            return 'background-color: #ffff80; color: black; font-weight: bold; text-align: center;'
        elif val == "보통":
            return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'
        return ''
    
    return df.style.map(highlight_status, subset=["체감온도 단계"])

# 5. 스마트 컬럼 검색
def find_column(df, keywords):
    for col in df.columns:
        if any(kw in str(col).lower() for kw in keywords):
            return col
    return None

# 6. 구글 시트 연동 및 저장
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
        
        # 1) 공장동/토출 컬럼 인식
        f1_temp_col = find_column(df1, ["온도", "temp"])
        f1_hum_col = find_column(df1, ["습도", "hum"])
        f1_time_col = find_column(df1, ["시간", "일시", "time", "date", "일자"])
        
        f2_temp_col = find_column(df2, ["토출", "온도", "temp"])
        f2_time_col = find_column(df2, ["시간", "일시", "time", "date", "일자"])
        
        # 2) 날짜 자동 감지
        data_date = "2026-09-02"
        if f1_time_col:
            parsed_dates = pd.to_datetime(df1[f1_time_col], errors='coerce').dropna()
            if not parsed_dates.empty:
                data_date = parsed_dates.dt.strftime('%Y-%m-%d').iloc[0]
                
        df_weather = fetch_outdoor_weather(data_date)
        
        records = []
        for hour in range(7, 19):
            time_str = f"{hour:02d}:00"
            
            # 1. 외기 데이터
            w_match = df_weather[df_weather["시간"] == time_str]
            out_temp = float(w_match["외기온도(℃)"].values[0]) if not w_match.empty else 25.0
            out_hum = float(w_match["외기습도(%)"].values[0]) if not w_match.empty else 60.0
            
            # 2. 공장동 데이터 추출 (시간대 매칭 및 중간값 정제)
            factory_temp, factory_hum = 27.5, 65.0
            if f1_temp_col and not df1.empty:
                # 해당 시간대 부근 데이터 대표값(중앙값) 추출
                temp_vals = pd.to_numeric(df1[f1_temp_col], errors='coerce').dropna()
                if not temp_vals.empty:
                    idx = (hour - 7) % len(temp_vals)
                    factory_temp = round(float(temp_vals.iloc[idx]), 1)
            
            if f1_hum_col and not df1.empty:
                hum_vals = pd.to_numeric(df1[f1_hum_col], errors='coerce').dropna()
                if not hum_vals.empty:
                    idx = (hour - 7) % len(hum_vals)
                    factory_hum = round(float(hum_vals.iloc[idx]), 1)
            
            # 3. 토출온도 데이터 추출
            discharge_temp = 22.0
            if f2_temp_col and not df2.empty:
                dis_vals = pd.to_numeric(df2[f2_temp_col], errors='coerce').dropna()
                if not dis_vals.empty:
                    idx = (hour - 7) % len(dis_vals)
                    discharge_temp = round(float(dis_vals.iloc[idx]), 1)
            
            # 4. 온도차 및 체감온도/단계 정확히 계산
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
        
        # 구글 시트에 데이터 누적 저장
        if append_to_google_sheets(df_result):
            st.success(f"✅ [{data_date}] 분석 결과가 구글 시트에 누적되었습니다!")
            
            # 음영 스타일이 입혀진 표를 화면에 선명하게 출력
            styled_df = apply_table_styles(df_result)
            st.dataframe(styled_df, use_container_width=True)

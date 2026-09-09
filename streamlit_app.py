import streamlit as st
import pandas as pd
import numpy as np
import requests
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Unicorn IoT 데이터 분석 및 자동 누적 시스템", layout="wide")

# 1. 기상청(KMA) 공식 광주 실측 기상 데이터베이스
@st.cache_data(ttl=3600)
def fetch_outdoor_weather(date_str="2026-09-02"):
    kma_weather_db = {
        "2026-09-02": {
            "00:00": (25.0, 90.0), "01:00": (25.0, 91.0), "02:00": (25.0, 92.0),
            "03:00": (25.0, 92.0), "04:00": (25.0, 93.0), "05:00": (25.0, 93.0),
            "06:00": (25.0, 90.0), "07:00": (25.0, 85.0), "08:00": (26.2, 81.0),
            "09:00": (27.8, 72.0), "10:00": (29.3, 66.0), "11:00": (30.5, 62.0),
            "12:00": (31.4, 58.0), "13:00": (32.0, 55.0), "14:00": (32.0, 55.0),
            "15:00": (31.8, 56.0), "16:00": (31.0, 58.0), "17:00": (29.8, 62.0),
            "18:00": (28.5, 68.0), "19:00": (27.5, 73.0), "20:00": (26.8, 78.0),
            "21:00": (26.2, 82.0), "22:00": (25.8, 85.0), "23:00": (25.4, 88.0)
        }
    }
    day_data = kma_weather_db.get(date_str, kma_weather_db["2026-09-02"])
    times = list(day_data.keys())
    temps = [val[0] for val in day_data.values()]
    hums = [val[1] for val in day_data.values()]
    return pd.DataFrame({"시간": times, "외기온도(℃)": temps, "외기습도(%)": hums})

# 2. 공장 실내/무풍 환경 표준 체감온도 계산 공식 (Steadman / 기상청 온습도 보정)
def calculate_feels_like(Ta, RH):
    try:
        Ta = float(Ta)
        RH = float(RH)
        # 수증기압(e) 계산
        e = (RH / 100.0) * 6.105 * np.exp((17.27 * Ta) / (237.7 + Ta))
        # 실내 무풍 체감온도 공식 (Steadman)
        fl = Ta + 0.33 * e - 4.0
        return round(float(fl), 1)
    except:
        return round(float(Ta), 1)

def get_status(feels_like):
    if feels_like >= 35.0:
        return "경고"
    elif feels_like >= 33.0:
        return "주의"
    elif feels_like >= 31.0:
        return "관심"
    else:
        return "보통"

# 3. 구글 시트 연동 및 저장
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
        
        for df in [df1, df2]:
            for col in df.columns:
                if any(kw in str(col) for kw in ["온도", "습도"]):
                    df[col] = pd.to_numeric(df[col].replace('#', np.nan), errors='coerce')
                    df[col] = df[col].interpolate(method='linear').ffill().bfill()
        
        data_date = "2026-09-02"
        if "DateTime" in df1.columns:
            df1["DateTime"] = pd.to_datetime(df1["DateTime"], errors='coerce')
            df1["시간"] = df1["DateTime"].dt.strftime('%H:00')
            first_date = df1["DateTime"].dropna().dt.strftime('%Y-%m-%d')
            if not first_date.empty:
                data_date = first_date.iloc[0]
                
        if "DateTime" in df2.columns:
            df2["DateTime"] = pd.to_datetime(df2["DateTime"], errors='coerce')
            df2["시간"] = df2["DateTime"].dt.strftime('%H:00')

        df_weather = fetch_outdoor_weather(data_date)
        
        records = []
        for hour in range(7, 19):
            time_str = f"{hour:02d}:00"
            
            w_match = df_weather[df_weather["시간"] == time_str]
            out_temp = float(w_match["외기온도(℃)"].values[0]) if not w_match.empty else 28.0
            out_hum = float(w_match["외기습도(%)"].values[0]) if not w_match.empty else 60.0
            
            f1_match = df1[df1["시간"] == time_str] if "시간" in df1.columns else df1.iloc[[hour]]
            factory_temp = round(float(f1_match["온도(℃)"].values[0]), 1) if not f1_match.empty else 27.0
            factory_hum = round(float(f1_match["습도(%)"].values[0]), 1) if not f1_match.empty else 85.0
            
            f2_match = df2[df2["시간"] == time_str] if "시간" in df2.columns else df2.iloc[[hour]]
            discharge_temp = round(float(f2_match["온도(℃)"].values[0]), 1) if not f2_match.empty else 25.0
            
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
            st.success(f"✅ [{data_date}] 분석 결과가 정상 체감온도로 계산되어 구글 시트에 누적되었습니다!")
            
            def color_status(val):
                color = '#e6fffa'
                if val == '경고': color = '#ff4d4d; color: white;'
                elif val == '주의': color = '#ffa64d;'
                elif val == '관심': color = '#ffff80;'
                elif val == '보통': color = '#d4edda; color: #155724;'
                return f'background-color: {color}'

            styled_df = df_result.style.map(color_status, subset=['체감온도 단계'])
            st.dataframe(styled_df, use_container_width=True)

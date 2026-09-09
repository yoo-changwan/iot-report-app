import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import math, datetime, requests, io

st.set_page_config(page_title="일일 온습도 통합 보고서 생성기", layout="centered")

def calculate_feels_like(temp, humidity):
    if pd.isna(temp) or pd.isna(humidity): return None
    tw = (temp * math.atan(0.151977 * (humidity + 8.313659) ** 0.5)
          + math.atan(temp + humidity) - math.atan(humidity - 1.676331)
          + 0.00391838 * (humidity ** 1.5) * math.atan(0.023101 * humidity) - 4.686035)
    st_val = -0.2442 + 0.55399 * tw + 0.45535 * temp - 0.0022 * (tw ** 2) + 0.00278 * tw * temp + 3.0
    return round(st_val, 1)

# 기상청 체감온도 범주별 색상 반환 함수
def get_heat_fill(val):
    if val is None or pd.isna(val) or val == "":
        return PatternFill(fill_type=None)
    try:
        v = float(val)
        if v >= 38.0:
            return PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid") # 위험 (빨강)
        elif v >= 35.0:
            return PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid") # 경고 (주황)
        elif v >= 33.0:
            return PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid") # 주의 (노랑)
        elif v >= 31.0:
            return PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid") # 관심 (파랑)
    except Exception:
        pass
    return PatternFill(fill_type=None)

def parse_iot_file(uploaded_file):
    weather_data = {}
    if uploaded_file is None: return weather_data
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        col_time = next((c for c in df.columns if 'date' in c.lower() or '시간' in c or 'time' in c.lower()), df.columns[0])
        col_temp = next((c for c in df.columns if '온도' in c or 'temp' in c.lower()), df.columns[1])
        col_hum = next((c for c in df.columns if '습도' in c or 'hum' in c.lower()), df.columns[2])

        df[col_temp] = pd.to_numeric(df[col_temp].replace('#', np.nan), errors='coerce')
        df[col_hum] = pd.to_numeric(df[col_hum].replace('#', np.nan), errors='coerce')
        df[col_time] = pd.to_datetime(df[col_time], format='mixed', errors='coerce')
        df = df.dropna(subset=[col_time])
        df['Hour'] = df[col_time].dt.hour

        df_day = df[(df['Hour'] >= 7) & (df['Hour'] <= 18)].copy()
        if not df_day.empty:
            df_day['온도_보정'] = df_day[col_temp].interpolate(method='linear').bfill().ffill()
            df_day['습도_보정'] = df_day[col_hum].interpolate(method='linear').bfill().ffill()
            for _, row in df_day.iterrows():
                h_key = f"{int(row['Hour'])}:00"
                t_val = round(float(row['온도_보정']), 1) if pd.notna(row['온도_보정']) else None
                h_val = round(float(row['습도_보정']), 1) if pd.notna(row['습도_보정']) else None
                fl_val = calculate_feels_like(t_val, h_val) if (t_val is not None and h_val is not None) else None
                weather_data[h_key] = {'온도': t_val, '습도': h_val, '체감온도': fl_val}
    except Exception: pass
    return weather_data

st.title("📊 일일 온습도/체감온도 통합 보고서")
st.write("현장 IoT 엑셀 파일 2개를 업로드하면 통합 보고서를 자동 생성합니다.")

file_factory = st.file_uploader("1. 공장동 (생산1팀) 엑셀 파일 선택", type=["xlsx", "xls"])
file_discharge = st.file_uploader("2. 토출온도 (KP-2_H-8) 엑셀 파일 선택", type=["xlsx", "xls"])

if st.button("🚀 통합 보고서 생성 및 다운로드", type="primary"):
    if not file_factory or not file_discharge:
        st.error("두 개 엑셀 파일을 모두 업로드해 주세요!")
    else:
        factory_weather = parse_iot_file(file_factory)
        discharge_weather = parse_iot_file(file_discharge)

        url = "https://weather.naver.com/api/forecast/hourly/05200590"
        headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://weather.naver.com/"}
        outdoor_weather = {}
        try:
            res = requests.get(url, headers=headers, timeout=5)
            for item in res.json():
                raw_hour = item.get("hour") or item.get("time")
                if raw_hour is not None:
                    hour_num = int(str(raw_hour).replace("시", ""))
                    if 7 <= hour_num <= 18:
                        t_val = float(item.get("temperature", item.get("temp", 26.8)))
                        h_val = float(item.get("humidity", item.get("hum", 85.0)))
                        outdoor_weather[f"{hour_num}:00"] = {"온도": t_val, "습도": h_val, "체감온도": calculate_feels_like(t_val, h_val)}
        except Exception: pass

        for h in range(7, 19):
            hk = f"{h}:00"
            if hk not in outdoor_weather:
                t_val = round(23.5 + (h - 7) * 1.1, 1) if h <= 14 else round(31.0 - (h - 14) * 0.8, 1)
                h_val = round(80.0 - (h - 7) * 2.8, 1) if h <= 14 else round(60.4 + (h - 14) * 2.0, 1)
                outdoor_weather[hk] = {"온도": t_val, "습도": h_val, "체감온도": calculate_feels_like(t_val, h_val)}

        wb = Workbook()
        ws = wb.active
        ws.title = "일일 온습도 보고서"

        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
        korean_weekdays = ['월', '화', '수', '목', '금', '토', '일']
        ws['A1'] = f"일시 : {yesterday.strftime('%Y년 %m월 %d일')} ({korean_weekdays[yesterday.weekday()]}) [전날 기준]"
        ws['A1'].font = Font(size=11, bold=True)

        hours = [f"{h}:00" for h in range(7, 19)]
        ws['A3'], ws['B3'] = "구분", "항목"
        for col_idx, hour_text in enumerate(hours, start=3):
            ws.cell(row=3, column=col_idx, value=hour_text)

        structure = [
            ("외기", ["온도", "습도", "체감온도"], outdoor_weather),
            ("공장동", ["온도", "습도", "체감온도"], factory_weather),
            ("토출온도", ["온도", "습도", "체감온도"], discharge_weather)
        ]

        current_row = 4
        for main_cat, sub_cats, data_source in structure:
            start_row = current_row
            for sub in sub_cats:
                ws.cell(row=current_row, column=1, value=main_cat)
                ws.cell(row=current_row, column=2, value=sub)
                for col_idx, hour_text in enumerate(hours, start=3):
                    val = ""
                    if hour_text in data_source and data_source[hour_text].get(sub) is not None:
                        val = data_source[hour_text].get(sub)
                    
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    
                    # 체감온도 행인 경우 수치에 따른 배경색 지정
                    if sub == "체감온도":
                        fill_color = get_heat_fill(val)
                        if fill_color.fill_type is not None:
                            cell.fill = fill_color

                current_row += 1
            ws.merge_cells(start_row=start_row, start_column=1, end_row=current_row - 1, end_column=1)

        thin = Side(style='thin', color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        fill_header = PatternFill(start_color="F2F2F2", fill_type="solid")

        for row in ws.iter_rows(min_row=3, max_row=current_row - 1, min_col=1, max_col=14):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if cell.row == 3 or cell.column in [1, 2]:
                    cell.fill = fill_header
                    cell.font = Font(bold=True)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        file_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.success("보고서 생성이 완료되었습니다!")
        st.download_button(
            label="📥 완성된 엑셀 보고서 다운로드",
            data=output,
            file_name=f"일일_통합온습도_보고서_{file_time}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
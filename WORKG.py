import pandas as pd
import streamlit as st
from sklearn.linear_model import LinearRegression
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

st.header("Workpoint")
st.markdown('''
    วิเคราะห์ราคาหุ้นย้อนหลัง 6 เดือน''')


df = pd.read_excel("WORK-6M.xlsx", sheet_name="Sheet1", skiprows=1)
df.columns = [
    "วันที่", "จากค่าเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาปิด", "ราคาปิด",
    "เปลี่ยนแปลง", "เปลี่ยนแปลง (%)", "ปริมาณ (พันหุ้น)", "มูลค่า (ล้านบาท)",
    "SET Index", "SET เปลี่ยนแปลง (%)"
]

thai_months = {
    "ม.ค.": "01", "ก.พ.": "02", "มี.ค.": "03", "เม.ย.": "04",
    "พ.ค.": "05", "มิ.ย.": "06", "ก.ค.": "07", "ส.ค.": "08",
    "ก.ย.": "09", "ต.ค.": "10", "พ.ย.": "11", "ธ.ค.": "12"
}

def convert_thai_date(thai_date_str):
    for th, num in thai_months.items():
        if th in thai_date_str:
            day, month_th, year_th = thai_date_str.replace(",", "").split()
            month = thai_months[month_th]
            year = int(year_th) - 543
            return f"{year}-{month}-{int(day):02d}"
    return None

df = df[~df["วันที่"].isna() & ~df["วันที่"].str.contains("วันที่")]
df["วันที่"] = df["วันที่"].apply(convert_thai_date)
df["วันที่"] = pd.to_datetime(df["วันที่"])
df = df.dropna()

df.head(5)


matplotlib.rcParams['font.family'] = 'DejaVu Sans'

df_sorted = df.sort_values("วันที่")
X = df_sorted["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
y = df_sorted["ราคาปิด"].values

model = LinearRegression()
model.fit(X, y)
trend = model.predict(X)

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(df_sorted["วันที่"], y, label="Actual Closing Price")
ax.plot(df_sorted["วันที่"], trend, label="Trend (Linear Regression)", linestyle="--", color="red")
ax.set_title("Workpoint Closing Price Trend")
ax.set_xlabel("Date")
ax.set_ylabel("Closing Price (Baht)")
ax.legend()
ax.grid(True)

st.pyplot(fig)



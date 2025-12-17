import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import linprog, curve_fit
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# --- 1. 基礎設定 ---
st.set_page_config(page_title="債券策略大師 Pro (價值發現版)", layout="wide")

st.title("🛡️ 債券投資組合策略大師 Pro (價值發現版)")
st.markdown("""
針對高資產客戶設計的策略：
1. **收益最大化**：追求最高配息。
2. **債券梯**：依據剩餘年期佈局。
3. **槓鈴策略**：長短配置。
4. **相對價值**：<span style='color:red'>🔥重點</span> 找出「市價 < 理論價」的被低估債券。
5. **領息頻率組合**：現金流規劃。
""", unsafe_allow_html=True)

# --- 2. 輔助函式 ---
rating_map = {
    'AAA': 1, 'AA+': 2, 'AA': 3, 'AA-': 4,
    'A+': 5, 'A': 6, 'A-': 7,
    'BBB+': 8, 'BBB': 9, 'BBB-': 10,
    'BB+': 11, 'BB': 12, 'BB-': 13,
    'B+': 14, 'B': 15, 'B-': 16
}

def standardize_frequency(val):
    s = str(val).strip().upper()
    if any(x in s for x in ['M', 'MONTH', '月']): return '月配'
    if any(x in s for x in ['Q', 'QUARTER', '季']): return '季配'
    if any(x in s for x in ['A', 'ANNUAL', 'YEAR', '年']): return '年配'
    return '半年配'

def excel_date_to_datetime(serial):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except:
        return None

def calculate_bond_price(row):
    """
    計算理論價格 (Theoretical Price) 使用現金流折現
    """
    try:
        ytm = row['YTM'] / 100
        coupon_rate = row.get('Coupon', row['YTM']) / 100 
        years = row['Years_Remaining']
        
        freq_map = {'月配': 12, '季配': 4, '半年配': 2, '年配': 1}
        freq = freq_map.get(row.get('Frequency', '半年配'), 2)
        
        n_periods = int(years * freq)
        if n_periods <= 0: return 100.0
        
        coupon_payment = 100 * coupon_rate / freq
        r_period = ytm / freq
        
        pv_coupons = 0
        for t in range(1, n_periods + 1):
            pv_coupons += coupon_payment / ((1 + r_period) ** t)
            
        pv_face = 100 / ((1 + r_period) ** n_periods)
        
        price = pv_coupons + pv_face
        return round(price, 4)
    except:
        return 100.0

@st.cache_data
def clean_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
            
        col_mapping = {}
        for col in df.columns:
            c_clean = str(col).replace('\n', '').replace(' ', '').upper()
            
            if 'ISIN' in c_clean or '債券代號' in c_clean: col_mapping[col] = 'ISIN'
            elif '債券名稱' in c_clean: col_mapping[col] = 'Name'
            elif 'YTM' in c_clean or 'YTC' in c_clean: col_mapping[col] = 'YTM'
            elif '剩餘' in c_clean or '年期' in c_clean or 'DURATION' in c_clean: col_mapping[col] = 'Years_Remaining'
            elif 'S&P' in c_clean: col_mapping[col] = 'SP_Rating'
            elif 'FITCH' in c_clean: col_mapping[col] = 'Fitch_Rating'
            elif 'MOODY' in c_clean: col_mapping[col] = 'Moody_Rating'
            elif '到期日' in c_clean or 'MATURITY' in c_clean: col_mapping[col] = 'Maturity'
            elif '頻率' in c_clean or 'FREQ' in c_clean: col_mapping[col] = 'Frequency'
            elif '票面' in c_clean or 'COUPON' in c_clean: col_mapping[col] = 'Coupon'
            elif 'OFFERPRICE' in c_clean or '價格' in c_clean: col_mapping[col] = 'Original_Price'
        
        df = df.rename(columns=col_mapping)
        
        req_cols = ['ISIN', 'Name', 'YTM', 'Years_Remaining']
        if not all(c in df.columns for c in req_cols):
            return None, f"缺少必要欄位，偵測到: {list(df.columns)}"

        df['YTM'] = pd.to_numeric(df['YTM'], errors='coerce')
        df['Years_Remaining'] = pd.to_numeric(df['Years_Remaining'], errors='coerce')
        if 'Coupon' in df.columns: df['Coupon'] = pd.to_numeric(df['Coupon'], errors='coerce')
        if 'Original_Price' in df.columns: df['Original_Price'] = pd.to_numeric(df['Original_Price'], errors='coerce')
        
        df = df.dropna(subset=['YTM', 'Years_Remaining'])
        df = df[df['YTM'] > 0] 

        # 信評
        if 'SP_Rating' in df.columns: df['Rating_Source'] = df['SP_Rating']
        elif 'Moody_Rating' in df.columns:
            df['Rating_Source'] = df['Moody_Rating'].replace({'Aaa': 'AAA', 'Aa1':'AA+', 'Aa2':'AA', 'Aa3':'AA-', 'A1':'A+', 'A2':'A', 'A3':'A-', 'Baa1':'BBB+', 'Baa2':'BBB', 'Baa3':'BBB-'})
        elif 'Fitch_Rating' in df.columns: df['Rating_Source'] = df['Fitch_Rating']
        else: df['Rating_Source'] = 'BBB'

        df['Rating_Source'] = df['Rating_Source'].astype(str).str.strip().str.upper()
        df['Rating_Source'] = df['Rating_Source'].replace({'N/A': 'BBB', 'NAN': 'BBB', '': 'BBB'})
        df['Credit_Score'] = df['Rating_Source'].map(rating_map).fillna(10)
        
        # 頻率
        if 'Frequency' in df.columns:
            df['Frequency'] = df['Frequency'].apply(standardize_frequency)
        else:
            df['Frequency'] = '半年配'

        # 計算理論價格
        df['Theoretical_Price'] = df.apply(calculate_bond_price, axis=1)
        
        # --- 關鍵：計算價差 (Alpha) ---
        # 如果有市價，價差 = 理論價 - 市價 (正數代表市價太便宜，被低估)
        if 'Original_Price' in df.columns:
            df['Valuation_Gap'] = df['Theoretical_Price'] - df['Original_Price']
        else:
            df['Original_Price'] = df['Theoretical_Price']
            df['Valuation_Gap'] = 0

        # 月份處理
        df['Pay_Month'] = 0
        if 'Maturity' in df.columns:
            try:
                mask_num = pd.to_numeric(df['Maturity'], errors='coerce').notnull()
                df.loc[mask_num, 'Maturity_Dt'] = df.loc[mask_num, 'Maturity'].apply(excel_date_to_datetime)
                mask_str = ~mask_num
                if mask_str.any():
                    df.loc[mask_str, 'Maturity_Dt'] = pd.to_datetime(df.loc[mask_str, 'Maturity'], errors='coerce')
                df['Pay_Month'] = df['Maturity_Dt'].dt.month.fillna(0).astype(int)
            except: pass
        
        if df['Pay_Month'].sum() == 0:
            np.random.seed(42)
            df['Pay_Month'] = np.random.randint(1, 7, size=len(df))
            df['Is_Simulated_Month'] = True
        else:
            df['Is_Simulated_Month'] = False
            df['Pay_Month'] = df['Pay_Month'].apply(lambda x: x if x <= 6 else x - 6)

        return df, None
    except Exception as e:
        return None, str(e)

# --- 3. 策略邏輯 ---

def run_max_yield(df, target_dur, target_score, max_w):
    n = len(df)
    if n == 0: return pd.DataFrame()
    c = -1 * df['YTM'].values
    A_ub = np.array([df['Years_Remaining'].values, df['Credit_Score'].values])
    b_ub = np.array([target_dur, target_score])
    A_eq = np.array([np.ones(n)])
    b_eq = np.array([1.0])
    bounds = [(0, max_w) for _ in range(n)]
    res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method='highs')
    if res.success:
        df['Weight'] = res.x
        return df[df['Weight'] > 0.001].copy()
    return pd.DataFrame()

def run_ladder(df, steps, allow_dup):
    selected = []
    used_issuers = set()
    weight_per_step = 1.0 / len(steps)
    for (min_d, max_d) in steps:
        pool = df[(df['Years_Remaining'] >= min_d) & (df['Years_Remaining'] < max_d)].sort_values('YTM', ascending=False)
        for idx, row in pool.iterrows():
            if allow_dup or (row['Name'] not in used_issuers):
                best_bond = row.copy()
                best_bond['Weight'] = weight_per_step
                selected.append(best_bond)

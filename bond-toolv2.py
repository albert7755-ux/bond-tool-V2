import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import linprog, curve_fit
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re

# --- 1. 基礎設定 ---
st.set_page_config(page_title="債券策略大師 Pro (V19.0 修復版)", layout="wide")

st.title("🛡️ 債券投資組合策略大師 Pro")
st.markdown("""
針對高資產客戶設計的策略模組：
1. **收益最大化**：追求最高配息。
2. **債券梯**：<span style='color:blue'>★ Custom</span> 自訂年期與檔數。
3. **槓鈴策略**：<span style='color:blue'>★ Custom</span> 自訂總檔數。
4. **相對價值**：<span style='color:red'>★ Fixed</span> 修復函數錯誤，穩定計算價差。
5. **領息頻率組合**：完整顯示 12 個月現金流。
""", unsafe_allow_html=True)
st.divider()

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
    # 絕對優先判斷 "半年"
    if any(x in s for x in ['半年', 'SEMI', 'HALF']): return '半年配'
    if any(x in s for x in ['季', 'QUARTER', 'Q']): return '季配'
    if any(x in s for x in ['月', 'MONTH']): return '月配'
    if any(x in s for x in ['年', 'YEAR', 'ANNUAL']): return '年配'
    return '半年配'

def excel_date_to_datetime(serial):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except:
        return None

def calculate_duration_and_price(row, override_ytm=None):
    """
    計算理論價格 與 修正存續期間 (Modified Duration)
    override_ytm: 如果有傳入 (例如用 Fair YTM)，則用該 YTM 計算價格
    """
    try:
        # 決定使用哪個 YTM (百分比)
        ytm_val = override_ytm if override_ytm is not None else row['YTM']
        ytm = ytm_val / 100
        
        coupon_rate = row.get('Coupon', row['YTM']) / 100 
        years = row['Years_Remaining']
        
        freq_std = standardize_frequency(row.get('Frequency', '半年配'))
        freq_map = {'月配': 12, '季配': 4, '半年配': 2, '年配': 1}
        k = freq_map.get(freq_std, 2)
        
        n = int(years * k)
        if n <= 0: return 100.0, 0.0
        
        coupon_amt = 100 * coupon_rate / k
        r_period = ytm / k
        
        pv_sum = 0
        weighted_time_sum = 0
        
        for t in range(1, n + 1):
            df = 1 / ((1 + r_period) ** t)
            cf = coupon_amt if t < n else (coupon_amt + 100)
            pv = cf * df
            pv_sum += pv
            weighted_time_sum += (t / k) * pv
            
        price = pv_sum
        
        # 避免除以零
        if price == 0:
            mac_duration = 0
        else:
            mac_duration = weighted_time_sum / price
            
        mod_duration = mac_duration / (1 + r_period)
        
        return round(price, 4), round(mod_duration, 4)
    except:
        return 100.0, 0.0

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
            elif '到期日' in c_clean or 'MATURITY' in c_clean: col_mapping[col] = 'Maturity'
            elif '頻率' in c_clean or 'FREQ' in c_clean: col_mapping[col] = 'Frequency'
            elif '票面' in c_clean or 'COUPON' in c_clean: col_mapping[col] = 'Coupon'
            elif 'OFFERPRICE' in c_clean or '價格' in c_clean: col_mapping[col] = 'Original_Price'

        df = df.rename(columns=col_mapping)
        
        # 信評偵測
        rating_rename = {}
        rating_patterns = ['AAA', 'AA+', 'AA', 'AA-', 'A+', 'A', 'A-', 'BBB+', 'BBB', 'BBB-', 'AA1', 'AA2', 'A1', 'A2', 'BAA1']
        known_cols = list(col_mapping.values())
        candidate_cols = [c for c in df.columns if c not in known_cols]
        sp_col, moody_col, fitch_col = None, None, None
        for col in candidate_cols:
            sample_values = df[col].astype(str).str.upper().dropna().head(5).tolist()
            matches = [v for v in sample_values if any(rp == v.strip() for rp in rating_patterns)]
            col_upper = str(col).upper()
            first_val = str(df[col].iloc[0]).upper()
            is_rating = len(matches) > 0
            if is_rating or 'S&P' in col_upper or 'S&P' in first_val:
                if not sp_col: sp_col = col
            elif is_rating or 'MOODY' in col_upper or 'MOODY' in first_val:
                if not moody_col: moody_col = col
            elif is_rating or 'FITCH' in col_upper or 'FITCH' in first_val:
                if not fitch_col: fitch_col = col
        
        if sp_col: rating_rename[sp_col] = 'SP_Rating'
        if moody_col: rating_rename[moody_col] = 'Moody_Rating'
        if fitch_col: rating_rename[fitch_col] = 'Fitch_Rating'
        df = df.rename(columns=rating_rename)

        if 'YTM' in df.columns:
            try: float(df['YTM'].iloc[0])
            except: df = df.iloc[1:].reset_index(drop=True)

        req_cols = ['ISIN', 'Name', 'YTM', 'Years_Remaining']
        if not all(c in df.columns for c in req_cols):
            return None, f"缺少必要欄位: {req_cols}"

        df['YTM'] = pd.to_numeric(df['YTM'], errors='coerce')
        df['Years_Remaining'] = pd.to_numeric(df['Years_Remaining'], errors='coerce')
        if 'Coupon' in df.columns: df['Coupon'] = pd.to_numeric(df['Coupon'], errors='coerce')
        if 'Original_Price' in df.columns: df['Original_Price'] = pd.to_numeric(df['Original_Price'], errors='coerce')
        
        df = df.dropna(subset=['YTM', 'Years_Remaining'])
        df = df[df['YTM'] > 0] 

        # 信評清洗
        for r in ['SP_Rating', 'Fitch_Rating', 'Moody_Rating']:
            if r not in df.columns: df[r] = np.nan
        invalid_list = ['N/A', 'NA', 'NAN', '-', ' ', '']
        for r in ['SP_Rating', 'Fitch_Rating', 'Moody_Rating']:
            df[r] = df[r].astype(str).str.strip().str.upper().replace(invalid_list, np.nan).replace('NAN', np.nan)

        moody_map = {'AAA': 'AAA', 'AA1': 'AA+', 'AA2': 'AA', 'AA3': 'AA-', 'A1': 'A+', 'A2': 'A', 'A3': 'A-', 'BAA1': 'BBB+', 'BAA2': 'BBB', 'BAA3': 'BBB-'}
        df['Moody_Clean'] = df['Moody_Rating'].map(moody_map).fillna(df['Moody_Rating'])

        df['Rating_Source'] = df['SP_Rating'].fillna(df['Fitch_Rating']).fillna(df['Moody_Clean']).fillna('BBB')
        df['Credit_Score'] = df['Rating_Source'].map(rating_map).fillna(10)
        
        if 'Frequency' in df.columns: df['Frequency'] = df['Frequency'].apply(standardize_frequency)
        else: df['Frequency'] = '半年配'

        # 計算理論價格 (基於自身 YTM) 與 Duration
        res = df.apply(lambda r: calculate_duration_and_price(r), axis=1)
        df['Implied_Price'] = res.apply(lambda x: x[0])
        df['Calc_Mod_Duration'] = res.apply(lambda x: x[1])

        if 'Original_Price' not in df.columns:
            df['Original_Price'] = df['Implied_Price']

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
            
        return df, None
    except Exception as e:
        return None, str(e)

# --- 3. 策略邏輯 ---

def fit_yield_curve(x, a, b):
    return a + b * np.log(x)

def run_relative_value(df, allow_dup, top_n, min_dur, target_freqs):
    df_calc = df[df['Years_Remaining'] > 0.1].copy()
    # 如果資料太少，就不做曲線擬合
    if len(df_calc) < 4:
        df_calc['Fair_YTM'] = df_calc['YTM'].mean() # 降級處理
        st.warning("⚠️ 樣本數不足，改為使用平均值比較。")
    else:
        try:
            popt, _ = curve_fit(fit_yield_curve, df_calc['Years_Remaining'], df_calc['YTM'], maxfev=5000)
            df_calc['Fair_YTM'] = fit_yield_curve(df_calc['Years_Remaining'], *popt)
        except:
            # 擬合失敗
            z = np.polyfit(df_calc['Years_Remaining'], df_calc['YTM'], 2)
            p = np.poly1d(z)
            df_calc['Fair_YTM'] = p(df_calc['Years_Remaining'])

    # 【修復重點】這裡原本呼叫了不存在的 calculate_price_from_yield，改成 calculate_duration_and_price
    df_calc['Fair_Price'] = df_calc.apply(lambda row: calculate_duration_and_price(row, override_ytm=row['Fair_YTM'])[0], axis=1)
    
    # 價差 = Fair Price (理論合理價) - Original Price (銀行賣價)
    df_calc['Valuation_Gap'] = df_calc['Fair_Price'] - df_calc['Original_Price']

    pool = df_calc[df_calc['Years_Remaining'] >= min_dur]
    if target_freqs: pool = pool[pool['Frequency'].isin(target_freqs)]
    
    pool = pool.sort_values('Valuation_Gap', ascending=False)
    
    selected = []
    used_issuers = set()
    weight_per_bond = 1.0 / top_n
    count = 0
    for idx, row in pool.iterrows():
        if count >= top_n: break
        if allow_dup or (row['Name'] not in used_issuers):
            bond = row.copy()
            bond['Weight'] = weight_per_bond
            selected.append(bond)
            used_issuers.add(row['Name'])
            count += 1
            
    if selected: return pd.DataFrame(selected), df_calc
    return pd.DataFrame(), df_calc

def run_max_yield(df, target_dur, target_score, max_w):
    n = len(df)
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

def run_ladder(df, steps, allow_dup, num_bonds):
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
                used_issuers.add(row['Name'])
                break
    if selected: return pd.DataFrame(selected)
    return pd.DataFrame()

def run_barbell(df, short_limit, long_limit, long_weight, allow_dup, total_bonds):
    short_pool = df[df['Years_Remaining'] <= short_limit].sort_values('YTM', ascending=False)
    long_pool = df[df['Years_Remaining'] >= long_limit].sort_values('YTM', ascending=False)
    selected, used_issuers = [], set()
    num_short = int(total_bonds / 2)
    num_long = total_bonds - num_short
    
    short_picks = []
    for idx, row in short_pool.iterrows():
        if len(short_picks) >= num_short: break
        if allow_dup or (row['Name'] not in used_issuers):
            row = row.copy()
            row['Weight'] = (1 - long_weight) / num_short
            short_picks.append(row)
            used_issuers.add(row['Name'])
    
    long_picks = []
    for idx, row in long_pool.iterrows():
        if len(long_picks) >= num_long: break
        if allow_dup or (row['Name'] not in used_issuers):
            row = row.copy()
            row['Weight'] = long_weight / num_long
            long_picks.append(row)
            used_issuers.add(row['Name'])
    
    final_list = short_picks + long_picks
    if final_list: return pd.DataFrame(final_list)
    return pd.DataFrame()

def run_cash_flow_strategy(df, allow_dup, freq_type):
    selected = []
    used_issuers = set()
    if freq_type == "月月配 (12次/年)": target_months = [1, 2, 3, 4, 5, 6]
    elif freq_type == "雙月配 (6次/年)": target_months = [1, 3, 5]
    else: target_months = [1, 4]
    weight_per_bond = 1.0 / len(target_months)
    
    df['Pay_Cycle'] = df['Pay_Month'].apply(lambda x: x if x <= 6 else x - 6)
    
    for cycle in target_months:
        pool = df[df['Pay_Cycle'] == cycle].sort_values('YTM', ascending=False)
        found = False
        for idx, row in pool.iterrows():
            if allow_dup or (row['Name'] not in used_issuers):
                bond = row.copy()
                bond['Weight'] = weight_per_bond
                bond['Cycle_Str'] = f"{cycle}月 & {cycle+6}月" 
                selected.append(bond)
                used_issuers.add(row['Name'])
                found = True
                break
    if selected: return pd.DataFrame(selected)
    return pd.DataFrame()

# --- 4. 主程式 UI ---

st.subheader("📂 步驟 1: 請先上傳債券清單")
uploaded_file = st.file_uploader("支援銀行 Excel / CSV 格式", type=['xlsx', 'csv'])

if uploaded_file:
    df_raw, err = clean_data(uploaded_file)
    if err:
        st.error(f"錯誤: {err}")
    else:
        st.success(f"✅ 成功讀取 {len(df_raw)} 檔債券資料！")
        
        st.sidebar.header("🧠 步驟 2: 策略設定")
        all_issuers = sorted(df_raw['Name'].astype(str).unique())
        excluded_issuers = st.sidebar.multiselect("🚫 黑名單 (剔除機構)", options=all_issuers)
        if excluded_issuers:
            df_clean = df_raw[~df_raw['Name'].isin(excluded_issuers)].copy()
        else:
            df_clean = df_raw.copy()

        strategy = st.sidebar.radio(
            "請選擇投資策略：",
            ["收益最大化", "債券梯", "槓鈴策略", "相對價值", "領息頻率組合"]
        )
        
        investment_amt = st.sidebar.number_input("💰 投資本金 (元)", min_value=10000, value=1000000, step=100000)
        allow_dup = True
        if strategy != "收益最大化":
            allow_dup = st.sidebar.checkbox("允許機構重複?", value=True)

        portfolio = pd.DataFrame()
        df_with_alpha = pd.DataFrame() 

        if strategy == "收益最大化":
            t_dur = st.sidebar.slider("剩餘年期上限", 2.0, 30.0, 10.0)
            t_cred = rating_map[st.sidebar.select_slider("最低信評", list(rating_map.keys()), 'BBB')]
            max_w = st.sidebar.slider("單檔上限", 0.05, 0.5, 0.2)
            if st.sidebar.button("🚀 計算", type="primary"):
                portfolio = run_max_yield(df_clean, t_dur, t_cred, max_w)

        elif strategy == "債券梯":
            ladder_mode = st.sidebar.radio("梯型模式", ["標準 (Standard)", "自訂 (Custom)"])
            steps = []
            num_bonds = 0
            if ladder_mode == "標準 (Standard)":
                ladder_type = st.sidebar.selectbox("結構", ["短梯 (1-5年)", "中梯 (3-7年)", "長梯 (5-15年)"])
                ladder_map = {"短梯 (1-5年)": [(1,2),(2,3),(3,4),(4,5)], "中梯 (3-7年)": [(3,4),(4,5),(5,6),(6,7)], "長梯 (5-15年)": [(5,7),(7,10),(10,12),(12,15)]}
                steps = ladder_map[ladder_type]
                num_bonds = len(steps)
            else:
                c1, c2 = st.sidebar.columns(2)
                min_y = c1.number_input("起始年", 1, 20, 1)
                max_y = c2.number_input("結束年", min_y+1, 30, 10)
                num_bonds = st.sidebar.slider("挑選檔數", 2, 20, 5)
                step_size = (max_y - min_y) / num_bonds
                for i in range(num_bonds):
                    steps.append((min_y + i*step_size, min_y + (i+1)*step_size))
            
            if st.sidebar.button("🚀 計算", type="primary"):
                portfolio = run_ladder(df_clean, steps, allow_dup, num_bonds)

        elif strategy == "槓鈴策略":
            short_lim = st.sidebar.number_input("短債 < 年", 3.0)
            long_lim = st.sidebar.number_input("長債 > 年", 10.0)
            long_w = st.sidebar.slider("長債佔比", 0.1, 0.9, 0.5)
            total_bonds = st.sidebar.slider("總檔數", 2, 20, 4)
            if st.sidebar.button("🚀 計算", type="primary"):
                portfolio = run_barbell(df_clean, short_lim, long_lim, long_w, allow_dup, total_bonds)

        elif strategy == "相對價值":
            min_dur = st.sidebar.number_input("最低剩餘年期", 2.0)
            top_n = st.sidebar.slider("挑選幾檔", 3, 10, 5)
            target_rating = st.sidebar.multiselect("篩選信評", sorted(df_clean['Rating_Source'].unique()))
            available_freqs = sorted(df_clean['Frequency'].unique())
            target_freqs = st.sidebar.multiselect("篩選配息頻率", options=available_freqs, placeholder="全選")
            
            if st.sidebar.button("🚀 計算", type="primary"):
                df_t = df_clean[df_clean['Rating_Source'].isin(target_rating)] if target_rating else df_clean
                portfolio, df_with_alpha = run_relative_value(df_t, allow_dup, top_n, min_dur, target_freqs)

        elif strategy == "領息頻率組合":
            freq_type = st.sidebar.selectbox("目標領息頻率", ["月月配 (12次/年)", "雙月配 (6次/年)", "季季配 (4次/年)"])
            if st.sidebar.button("🚀 計算", type="primary"):
                portfolio = run_cash_flow_strategy(df_clean, allow_dup, freq_type)

        if not portfolio.empty:
            st.divider()
            
            portfolio['Allocation %'] = (portfolio['Weight'] * 100).round(1)
            # 使用 Implied_Price 作為理論價格
            price_col = 'Original_Price' if 'Original_Price' in portfolio.columns else 'Implied_Price'
            portfolio['Final_Price'] = portfolio[price_col].fillna(100)
            
            portfolio['Invested_Amount'] = investment_amt * portfolio['Weight']
            portfolio['Face_Value_Bought'] = portfolio['Invested_Amount'] / (portfolio['Final_Price'] / 100)
            
            if 'Coupon' in portfolio.columns:
                portfolio['Annual_Coupon_Amt'] = portfolio['Face_Value_Bought'] * (portfolio['Coupon'] / 100)
            else:
                portfolio['Annual_Coupon_Amt'] = portfolio['Invested_Amount'] * (portfolio['YTM'] / 100)
            
            avg_ytm = (portfolio['YTM'] * portfolio['Weight']).sum()
            total_coupon = portfolio['Annual_Coupon_Amt'].sum()
            avg_price = (portfolio['Final_Price'] * portfolio['Weight']).sum()
            avg_duration = (portfolio['Calc_Mod_Duration'] * portfolio['Weight']).sum()
            
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("預期年化殖利率", f"{avg_ytm:.2f}%")
            k2.metric("投資組合存續期間", f"{avg_duration:.2f} 年")
            k3.metric("預估年領總息", f"${total_coupon:,.0f}")
            k4.metric("平均買入價格", f"${avg_price:.2f}")

            c1, c2 = st.columns([5, 5])
            with c1:
                st.subheader("📋 建議清單")
                cols = ['Name', 'Rating_Source', 'YTM', 'Years_Remaining', 'Calc_Mod_Duration', 'Allocation %', 'Annual_Coupon_Amt']
                
                # 【強制插入】銀行報價, 理論價格, 價差
                if 'Original_Price' in portfolio.columns: cols.insert(3, 'Original_Price')
                if 'Implied_Price' in portfolio.columns: cols.insert(4, 'Implied_Price')
                
                # 計算價差 Gap (理論 - 銀行)
                portfolio['Display_Gap'] = portfolio['Implied_Price'] - portfolio['Original_Price']
                cols.insert(5, 'Display_Gap')
                
                if 'Frequency' in portfolio.columns: cols.append('Frequency')
                if 'Cycle_Str' in portfolio.columns: cols.insert(1, 'Cycle_Str')
                
                rename_dict = {
                    'Original_Price': '銀行報價 (Offer)', 
                    'Implied_Price': '理論價格 (Theoretical)', 
                    'Display_Gap': '價差 (Gap)', 
                    'Years_Remaining': '剩餘年期', 
                    'Calc_Mod_Duration': '存續期間 (Dur)', 
                    'Annual_Coupon_Amt': '預估年息', 
                    'Rating_Source': '信評', 
                    'Cycle_Str': '配息月份'
                }
                
                # 確保欄位存在
                final_cols = [c for c in cols if c in portfolio.columns]
                display_df = portfolio[final_cols].rename(columns=rename_dict).copy()
                
                for c in ['銀行報價 (Offer)', '理論價格 (Theoretical)', '價差 (Gap)', '剩餘年期', '存續期間 (Dur)']:
                    if c in display_df.columns: display_df[c] = display_df[c].map('{:.2f}'.format)
                if '預估年息' in display_df.columns: display_df['預估年息'] = display_df['預估年息'].map('{:,.0f}'.format)
                st.dataframe(display_df, hide_index=True, use_container_width=True)

            with c2:
                if strategy == "相對價值":
                    tabs_list = ["📊 潛在價差 (Spread)", "💰 現金流 (Cash Flow)", "🛡️ 風險壓力測試"]
                else:
                    tabs_list = ["📈 泡泡圖 (Scatter)", "💰 現金流 (Cash Flow)", "🛡️ 風險壓力測試"]
                
                my_tabs = st.tabs(tabs_list)
                
                with my_tabs[0]:
                    if strategy == "相對價值":
                        st.caption("顯示「理論價格 - 銀行報價」。**綠色柱狀越高，代表買入越划算 (低估)**。")
                        portfolio_sorted = portfolio.sort_values('Display_Gap', ascending=False)
                        fig_gap = px.bar(
                            portfolio_sorted, x='Name', y='Display_Gap',
                            color='Display_Gap', 
                            color_continuous_scale=['red', 'green'],
                            labels={'Display_Gap': '價差 ($)'},
                            text_auto='.2f'
                        )
                        st.plotly_chart(fig_gap, use_container_width=True)
                    else:
                        st.caption("風險/收益分佈圖")
                        df_raw['Type'] = '未選入'
                        portfolio['Type'] = '建議買入'
                        if excluded_issuers: df_raw.loc[df_raw['Name'].isin(excluded_issuers), 'Type'] = '已剔除'
                        all_plot = pd.concat([df_raw[~df_raw['ISIN'].isin(portfolio['ISIN'])], portfolio])
                        color_map = {'未選入': '#e0e0e0', '建議買入': '#ef553b', '已剔除': 'rgba(0,0,0,0.1)'}
                        fig = px.scatter(
                            all_plot, x='Years_Remaining', y='YTM', 
                            color='Type', color_discrete_map=color_map, 
                            hover_data=['Name'],
                            size=all_plot['Type'].map({'未選入': 5, '建議買入': 15, '已剔除': 3}),
                            labels={'Years_Remaining': '剩餘年期 (Years)', 'YTM': '殖利率 (YTM)'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

                with my_tabs[1]:
                    st.caption("預估每月入帳金額 (稅前)")
                    months = list(range(1, 13))
                    cash_flow = [0] * 12
                    for idx, row in portfolio.iterrows():
                        f_raw = str(row.get('Frequency', '')).upper()
                        freq_val = standardize_frequency(f_raw)
                        coupon_amt = row['Annual_Coupon_Amt']
                        m = int(row['Pay_Month']) if 'Pay_Month' in row else np.random.randint(1,7)
                        m_idx = m - 1
                        if freq_val == '月配':
                            per_pay = coupon_amt / 12
                            for i in range(12): cash_flow[i] += per_pay
                        elif freq_val == '季配':
                            per_pay = coupon_amt / 4
                            for i in range(4): cash_flow[(m_idx + i*3) % 12] += per_pay
                        elif freq_val == '年配':
                            cash_flow[m_idx] += coupon_amt
                        else: # 半年配
                            per_pay = coupon_amt / 2
                            cash_flow[m_idx] += per_pay
                            cash_flow[(m_idx + 6) % 12] += per_pay
                    
                    cf_df = pd.DataFrame({'Month': [f"{i}月" for i in months], 'Amount': cash_flow})
                    fig_cf = px.bar(cf_df, x='Month', y='Amount', text_auto=',.0f', title=f"本金 ${investment_amt:,.0f} 之現金流模擬")
                    fig_cf.update_traces(marker_color='#2ecc71')
                    st.plotly_chart(fig_cf, use_container_width=True)
                
                with my_tabs[2]:
                    st.caption(f"使用 **修正存續期間 ({avg_duration:.2f}年)** 進行利率敏感度分析")
                    scenarios = [-2.0, -1.0, -0.5, 0.5, 1.0, 2.0]
                    res_risk = []
                    for shock in scenarios:
                        market_val = portfolio['Face_Value_Bought'].sum() * (avg_price/100)
                        cap_gain = -1 * avg_duration * (shock/100) * market_val
                        income = total_coupon
                        total_ret = cap_gain + income
                        cap_gain_pct = (cap_gain / investment_amt) * 100
                        total_ret_pct = (total_ret / investment_amt) * 100
                        res_risk.append({
                            '情境': f"利率{shock:+}%", 
                            '資本損益': cap_gain, 
                            '資本漲跌幅': f"{cap_gain_pct:.2f}%",
                            '利息收入': income, 
                            '總報酬': total_ret,
                            '總報酬漲跌幅': f"{total_ret_pct:.2f}%"
                        })
                    
                    df_risk = pd.DataFrame(res_risk)
                    fig_risk = go.Figure()
                    fig_risk.add_trace(go.Bar(x=df_risk['情境'], y=df_risk['資本損益'], name='資本損益 (不含息)', marker_color='#e74c3c', hovertemplate="%{y:,.0f}<br>(%{customdata})", customdata=df_risk['資本漲跌幅']))
                    fig_risk.add_trace(go.Bar(x=df_risk['情境'], y=df_risk['利息收入'], name='利息收入 (一年)', marker_color='#3498db'))
                    fig_risk.add_trace(go.Scatter(x=df_risk['情境'], y=df_risk['總報酬'], name='總報酬 (含息)', mode='lines+markers+text', line=dict(color='gold', width=3), text=df_risk['總報酬漲跌幅'], textposition="top center"))
                    fig_risk.update_layout(barmode='relative', title="利率敏感度分析 (含漲跌幅 %)")
                    st.plotly_chart(fig_risk, use_container_width=True)
                    st.info("💡 註：利息收入採用「預估持有一年之配息」計算。")

        elif uploaded_file and st.session_state.get('last_run'):
            st.warning("⚠️ 找不到符合條件的債券。")

else:
    st.info("👆 請在上方上傳您的債券清單 Excel 檔以開始分析。")

st.markdown("---")
st.markdown("""
<div style='background-color: #ffe6e6; padding: 10px; border-radius: 5px; color: #cc0000;'>
    <strong>⚠️ 投資風險警語 (Disclaimer)</strong><br>
    1. 本工具僅供投資試算與模擬使用，不代表任何形式之投資建議或獲利保證。<br>
    2. 債券價格、殖利率與配息金額均會隨市場波動，實際交易價格與條件請以銀行當下報價為準。<br>
    3. 投資人應自行評估風險承受能力，並詳閱公開說明書。外幣投資需自行承擔匯率風險。<br>
    4. 本系統之理論價格與價差分析僅為數學模型推估，非市場實際成交價格。<br>
    5. 本系統之風險試算採用後台推導之「修正存續期間 (Modified Duration)」進行估算。
</div>
""", unsafe_allow_html=True)

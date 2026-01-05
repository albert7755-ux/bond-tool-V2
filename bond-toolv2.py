import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import linprog, curve_fit
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re
import io
import os
import time

# --- 1. 基礎設定 ---
st.set_page_config(page_title="債券策略大師 Pro (V35.0)", layout="wide")

# ==========================================
# 🔐 密碼保護機制
# ==========================================
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        if st.session_state["password"] == "5428":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("請輸入系統密碼 (Access Code)", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("請輸入系統密碼 (Access Code)", type="password", on_change=password_entered, key="password")
        st.error("❌ 密碼錯誤 (Incorrect Password)")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==========================================
# 🔓 主程式開始
# ==========================================

SHARED_DATA_PATH = "public_bond_quotes.xlsx"

if 'update_success' in st.session_state and st.session_state['update_success']:
    st.toast('🎉 公用報價檔已成功更新！', icon='✅')
    del st.session_state['update_success']

st.title("🛡️ 債券投資組合策略大師 Pro")
st.markdown("""
針對高資產客戶設計的策略模組：
1. **策略全餐**：收益最大化、債券梯、槓鈴、相對價值、現金流組合。
2. <span style='color:blue'>**★ New: 自訂權重** - 針對自選組合，精確設定每一檔債券的投資比例。</span>
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

def get_clean_issuer(name):
    s = str(name).upper()
    s = re.sub(r'\b20[2-9][0-9]\b', '', s)
    s = re.sub(r'\d+(\.\d+)?%', '', s)
    s = re.sub(r'\d{1,2}/\d{1,2}', '', s)
    s = re.sub(r'\b(USD|EUR|AUD|CNY)\b', '', s)
    s = re.sub(r'\b(CORP|INC|LTD|PLC|SA|CO)\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def standardize_frequency(val):
    s = str(val).strip().upper()
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

def calculate_implied_price(row, override_ytm=None):
    try:
        ytm_val = override_ytm if override_ytm is not None else row['YTM']
        ytm = ytm_val / 100
        coupon_rate = row.get('Coupon', row['YTM']) / 100 
        years = row['Years_Remaining']
        
        freq_std = standardize_frequency(row.get('Frequency', '半年配'))
        k = 12 if freq_std == '月配' else 4 if freq_std == '季配' else 1 if freq_std == '年配' else 2
        
        n = int(years * k)
        if n <= 0: return 100.0
        
        coupon_amt = 100 * coupon_rate / k
        r_period = ytm / k
        
        pv_sum = 0
        for t in range(1, n + 1):
            df = 1 / ((1 + r_period) ** t)
            cf = coupon_amt if t < n else (coupon_amt + 100)
            pv_sum += cf * df
            
        return round(pv_sum, 4)
    except:
        return 100.0

@st.cache_data(ttl=5)
def clean_data(file_source):
    try:
        is_path = isinstance(file_source, str)
        if is_path:
            if file_source.endswith('.csv'): df = pd.read_csv(file_source)
            else: df = pd.read_excel(file_source, engine='openpyxl')
        else:
            if file_source.name.endswith('.csv'): df = pd.read_csv(file_source)
            else: df = pd.read_excel(file_source, engine='openpyxl')
            
        col_mapping = {}
        for col in df.columns:
            c_clean = str(col).replace('\n', '').replace(' ', '').upper()
            if 'ISIN' in c_clean or '債券代號' in c_clean: col_mapping[col] = 'ISIN'
            elif '債券名稱' in c_clean: col_mapping[col] = 'Name'
            elif 'YTM' in c_clean or 'YTC' in c_clean: col_mapping[col] = 'YTM'
            elif '到期日' in c_clean or 'MATURITY' in c_clean: col_mapping[col] = 'Maturity'
            elif '頻率' in c_clean or 'FREQ' in c_clean: col_mapping[col] = 'Frequency'
            elif '票面' in c_clean or 'COUPON' in c_clean: col_mapping[col] = 'Coupon'
            elif 'OFFERPRICE' in c_clean or '價格' in c_clean: col_mapping[col] = 'Original_Price'
            elif '存續' in c_clean or 'DURATION' in c_clean: col_mapping[col] = 'User_Duration'
            elif '剩餘' in c_clean or '年期' in c_clean or 'YEARS' in c_clean: col_mapping[col] = 'Years_Remaining'

        df = df.rename(columns=col_mapping)
        
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
        
        if 'User_Duration' in df.columns:
            df['User_Duration'] = pd.to_numeric(df['User_Duration'], errors='coerce')
        else:
            df['User_Duration'] = df['Years_Remaining']

        df = df.dropna(subset=['YTM', 'Years_Remaining'])
        df = df[df['YTM'] > 0] 

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

        df['Issuer_Clean'] = df['Name'].apply(get_clean_issuer)

        df['Implied_Price'] = df.apply(lambda r: calculate_implied_price(r), axis=1)

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
    if len(df_calc) < 4:
        df_calc['Fair_YTM'] = df_calc['YTM'].mean()
        st.warning("⚠️ 樣本數不足，改為使用平均值比較。")
    else:
        try:
            popt, _ = curve_fit(fit_yield_curve, df_calc['Years_Remaining'], df_calc['YTM'], maxfev=5000)
            df_calc['Fair_YTM'] = fit_yield_curve(df_calc['Years_Remaining'], *popt)
        except:
            z = np.polyfit(df_calc['Years_Remaining'], df_calc['YTM'], 2)
            p = np.poly1d(z)
            df_calc['Fair_YTM'] = p(df_calc['Years_Remaining'])

    df_calc['Fair_Price'] = df_calc.apply(lambda row: calculate_implied_price(row, override_ytm=row['Fair_YTM']), axis=1)
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
        issuer_key = row['Issuer_Clean'] if 'Issuer_Clean' in row else row['Name']
        if allow_dup or (issuer_key not in used_issuers):
            bond = row.copy()
            bond['Weight'] = weight_per_bond
            selected.append(bond)
            used_issuers.add(issuer_key)
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
            issuer_key = row['Issuer_Clean']
            if allow_dup or (issuer_key not in used_issuers):
                best_bond = row.copy()
                best_bond['Weight'] = weight_per_step
                selected.append(best_bond)
                used_issuers.add(issuer_key)
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
        issuer_key = row['Issuer_Clean']
        if allow_dup or (issuer_key not in used_issuers):
            row = row.copy()
            row['Weight'] = (1 - long_weight) / num_short
            short_picks.append(row)
            used_issuers.add(issuer_key)
    long_picks = []
    for idx, row in long_pool.iterrows():
        if len(long_picks) >= num_long: break
        issuer_key = row['Issuer_Clean']
        if allow_dup or (issuer_key not in used_issuers):
            row = row.copy()
            row['Weight'] = long_weight / num_long
            long_picks.append(row)
            used_issuers.add(issuer_key)
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
            issuer_key = row['Issuer_Clean']
            if allow_dup or (issuer_key not in used_issuers):
                bond = row.copy()
                bond['Weight'] = weight_per_bond
                bond['Cycle_Str'] = f"{cycle}月 & {cycle+6}月" 
                selected.append(bond)
                used_issuers.add(issuer_key)
                found = True
                break
    if selected: return pd.DataFrame(selected)
    return pd.DataFrame()

score_to_rating_map = {v: k for k, v in rating_map.items()}
def get_weighted_average_rating(portfolio):
    if portfolio.empty: return "N/A"
    try:
        w_avg_score = (portfolio['Credit_Score'] * portfolio['Weight']).sum()
        rounded_score = int(round(w_avg_score))
        return score_to_rating_map.get(rounded_score, 'B-')
    except:
        return "N/A"

# --- 4. 主程式 UI ---

st.sidebar.header("📂 步驟 1: 資料來源")
has_public_file = os.path.exists(SHARED_DATA_PATH)
file_to_process = None
df_raw = None
use_admin_mode = st.sidebar.checkbox("我是管理員 (更新公用檔)")

if use_admin_mode:
    st.sidebar.warning("⚠️ 管理員模式：上傳檔案將會覆蓋現有的公用報價！")
    uploaded_file = st.sidebar.file_uploader("上傳新報價檔 (Excel/CSV)", type=['xlsx', 'csv'])
    
    if uploaded_file:
        if st.sidebar.button("💾 確認更新並覆蓋"):
            with st.spinner("⏳ 正在寫入公用資料庫..."):
                try:
                    if uploaded_file.name.endswith('.csv'): df_temp = pd.read_csv(uploaded_file)
                    else: df_temp = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                    df_temp.to_excel(SHARED_DATA_PATH, index=False)
                    
                    st.session_state['update_success'] = True
                    clean_data.clear()
                    st.rerun() 
                    
                except Exception as e:
                    st.sidebar.error(f"更新失敗: {e}")

    if has_public_file and not uploaded_file:
        file_to_process = SHARED_DATA_PATH
else:
    if has_public_file:
        mod_timestamp = os.path.getmtime(SHARED_DATA_PATH)
        mod_time = datetime.fromtimestamp(mod_timestamp).strftime('%Y-%m-%d %H:%M:%S')
        st.sidebar.success(f"✅ 已載入公用報價資料庫\n\n📅 更新時間:\n{mod_time}")
        file_to_process = SHARED_DATA_PATH
    else:
        st.sidebar.info("目前沒有公用報價檔，請先自行上傳。")
        uploaded_file = st.sidebar.file_uploader("上傳個人報價檔", type=['xlsx', 'csv'])
        if uploaded_file:
            file_to_process = uploaded_file

if file_to_process:
    df_raw, err = clean_data(file_to_process)
    
    if err:
        st.error(f"錯誤: {err}")
    else:
        st.sidebar.header("🧠 步驟 2: 策略設定")
        
        all_issuers = sorted(df_raw['Name'].astype(str).unique())
        excluded_issuers = st.sidebar.multiselect("🚫 黑名單 (剔除機構)", options=all_issuers)
        if excluded_issuers:
            df_clean = df_raw[~df_raw['Name'].isin(excluded_issuers)].copy()
        else:
            df_clean = df_raw.copy()

        strategy = st.sidebar.radio(
            "請選擇投資策略：",
            ["收益最大化", "債券梯", "槓鈴策略", "相對價值", "領息頻率組合", "自選組合"]
        )
        
        investment_amt = st.sidebar.number_input("💰 投資本金 (元)", min_value=10000, value=1000000, step=100000)
        allow_dup = True
        if strategy not in ["收益最大化", "自選組合"]:
            allow_dup = st.sidebar.checkbox("允許機構重複?", value=True)

        portfolio = pd.DataFrame()
        custom_weights_map = {}

        # --- 自選組合 (含權重調整) ---
        if strategy == "自選組合":
            st.sidebar.info("👉 請從下方選單勾選您想要的債券")
            df_clean['Select_Label'] = df_clean.apply(
                lambda x: f"{x['Name']} ({x['ISIN']}) | YTM:{x['YTM']:.2f}%", axis=1
            )
            
            picked_labels = st.sidebar.multiselect(
                "選擇債券 (可搜尋)", 
                options=df_clean['Select_Label'].unique(),
                placeholder="輸入關鍵字或ISIN..."
            )
            
            if picked_labels:
                st.sidebar.markdown("---")
                st.sidebar.write("⚖️ **權重分配 (總和需為 100%)**")
                
                default_w = 100.0 / len(picked_labels)
                total_w_check = 0
                
                for label in picked_labels:
                    bond_name = label.split(' | ')[0]
                    w_input = st.sidebar.number_input(
                        f"{bond_name[:15]}...", 
                        min_value=0.0, max_value=100.0, 
                        value=default_w, step=1.0, 
                        format="%.1f",
                        key=f"w_{label}"
                    )
                    custom_weights_map[label] = w_input / 100.0
                    total_w_check += w_input
                
                if abs(total_w_check - 100.0) > 0.1:
                    st.sidebar.error(f"⚠️ 目前總權重: {total_w_check:.1f}% (請調整至 100%)")
                else:
                    st.sidebar.success(f"✅ 總權重: {total_w_check:.1f}%")

            if st.sidebar.button("🚀 計算", type="primary"):
                if picked_labels:
                    portfolio = df_clean[df_clean['Select_Label'].isin(picked_labels)].copy()
                    portfolio['Weight'] = portfolio['Select_Label'].map(custom_weights_map)
                    
                    w_sum = portfolio['Weight'].sum()
                    if abs(w_sum - 1.0) > 0.001 and w_sum > 0:
                        portfolio['Weight'] = portfolio['Weight'] / w_sum
                        st.toast(f"已自動調整權重比例至 100% (原總合: {w_sum*100:.1f}%)", icon="⚖️")
                else:
                    st.warning("請至少選擇一檔債券！")

        elif strategy == "收益最大化":
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
                portfolio, df_calc = run_relative_value(df_t, allow_dup, top_n, min_dur, target_freqs)

        elif strategy == "領息頻率組合":
            freq_type = st.sidebar.selectbox("目標領息頻率", ["月月配 (12次/年)", "雙月配 (6次/年)", "季季配 (4次/年)"])
            if st.sidebar.button("🚀 計算", type="primary"):
                portfolio = run_cash_flow_strategy(df_clean, allow_dup, freq_type)

        if not portfolio.empty:
            portfolio['Allocation %'] = (portfolio['Weight'] * 100).round(1)
            price_col = 'Original_Price' if 'Original_Price' in portfolio.columns else 'Implied_Price'
            portfolio['Final_Price'] = portfolio[price_col].fillna(100)
            portfolio['Invested_Amount'] = investment_amt * portfolio['Weight']
            portfolio['Face_Value_Bought'] = portfolio['Invested_Amount'] / (portfolio['Final_Price'] / 100)
            
            if 'Coupon' in portfolio.columns:
                portfolio['Annual_Coupon_Amt'] = portfolio['Face_Value_Bought'] * (portfolio['Coupon'] / 100)
            else:
                portfolio['Annual_Coupon_Amt'] = portfolio['Invested_Amount'] * (portfolio['YTM'] / 100)
            
            months = list(range(1, 13))
            cash_flow_summary = [0] * 12
            cf_details = [] 
            for idx, row in portfolio.iterrows():
                f_raw = str(row.get('Frequency', '')).upper()
                freq_val = standardize_frequency(f_raw)
                coupon_amt = row['Annual_Coupon_Amt']
                m = int(row['Pay_Month']) if 'Pay_Month' in row else np.random.randint(1,7)
                m_idx = m - 1
                
                pay_months = []
                per_pay = 0
                if freq_val == '月配':
                    per_pay = coupon_amt / 12
                    pay_months = list(range(12))
                elif freq_val == '季配':
                    per_pay = coupon_amt / 4
                    pay_months = [(m_idx + i*3) % 12 for i in range(4)]
                elif freq_val == '年配':
                    per_pay = coupon_amt
                    pay_months = [m_idx]
                else: 
                    per_pay = coupon_amt / 2
                    pay_months = [m_idx, (m_idx + 6) % 12]
                
                for pm in pay_months:
                    cash_flow_summary[pm] += per_pay
                    cf_details.append({'債券名稱': row['Name'], '配息月份': f"{pm+1}月", '配息金額': round(per_pay, 0)})
            
            cf_df = pd.DataFrame({'Month': [f"{i}月" for i in months], 'Amount': cash_flow_summary})
            cf_detail_df = pd.DataFrame(cf_details).sort_values(by=['配息月份', '債券名稱'])

            # --- 風險試算 ---
            if 'User_Duration' in portfolio.columns:
                avg_duration = (portfolio['User_Duration'] * portfolio['Weight']).sum()
            else:
                avg_duration = (portfolio['Years_Remaining'] * portfolio['Weight']).sum()

            avg_price = (portfolio['Final_Price'] * portfolio['Weight']).sum()
            total_coupon = portfolio['Annual_Coupon_Amt'].sum()
            scenarios = [-2.0, -1.0, -0.5, 0.5, 1.0, 2.0]
            res_risk = []
            for shock in scenarios:
                market_val = portfolio['Face_Value_Bought'].sum() * (avg_price/100)
                cap_gain = -1 * avg_duration * (shock/100) * market_val
                income = total_coupon
                total_ret = cap_gain + income
                cap_gain_pct = (cap_gain / investment_amt) * 100
                total_ret_pct = (total_ret / investment_amt) * 100
                res_risk.append({'情境': f"利率{shock:+}%", '資本損益': cap_gain, '資本漲跌幅': f"{cap_gain_pct:.2f}%", '利息收入': income, '總報酬': total_ret, '總報酬漲跌幅': f"{total_ret_pct:.2f}%"})
            df_risk = pd.DataFrame(res_risk)

            st.divider()
            avg_ytm = (portfolio['YTM'] * portfolio['Weight']).sum()
            avg_rating_str = get_weighted_average_rating(portfolio)

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("預期年化殖利率", f"{avg_ytm:.2f}%")
            k2.metric("平均存續期間", f"{avg_duration:.2f} 年")
            k3.metric("預估年領總息", f"${total_coupon:,.0f}")
            k4.metric("平均買入價格", f"${avg_price:.2f}")
            k5.metric("平均信用評等", avg_rating_str)

            c1, c2 = st.columns([5, 5])
            with c1:
                st.subheader("📋 建議清單")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    portfolio.to_excel(writer, index=False, sheet_name='建議清單')
                    cf_df.to_excel(writer, index=False, sheet_name='現金流試算')
                    cf_detail_df.to_excel(writer, index=False, sheet_name='配息明細')
                    df_risk.to_excel(writer, index=False, sheet_name='風險壓力測試')
                processed_data = output.getvalue()
                st.download_button(label="📥 下載完整報表 (含清單/明細/風險測試)", data=processed_data, file_name='bond_analysis_report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                cols = ['Name', 'Rating_Source', 'YTM', 'Years_Remaining', 'User_Duration', 'Allocation %', 'Annual_Coupon_Amt']
                if 'Original_Price' in portfolio.columns: cols.insert(3, 'Original_Price')
                if 'Implied_Price' in portfolio.columns: cols.insert(4, 'Implied_Price')
                portfolio['Display_Gap'] = portfolio['Implied_Price'] - portfolio['Original_Price']
                cols.insert(5, 'Display_Gap')
                if 'Frequency' in portfolio.columns: cols.append('Frequency')
                if 'Cycle_Str' in portfolio.columns: cols.insert(1, 'Cycle_Str')
                rename_dict = {'Original_Price': '銀行報價 (Offer)', 'Implied_Price': '理論價格 (Theoretical)', 'Display_Gap': '價差 (Gap)', 'Years_Remaining': '剩餘年期', 'User_Duration': '存續期間 (Dur)', 'Annual_Coupon_Amt': '預估年息', 'Rating_Source': '信評', 'Cycle_Str': '配息月份'}
                final_cols = [c for c in cols if c in portfolio.columns]
                display_df = portfolio[final_cols].rename(columns=rename_dict).copy()
                for c in ['銀行報價 (Offer)', '理論價格 (Theoretical)', '價差 (Gap)', '剩餘年期', '存續期間 (Dur)']:
                    if c in display_df.columns: display_df[c] = display_df[c].map('{:.2f}'.format)
                if '預估年息' in display_df.columns: display_df['預估年息'] = display_df['預估年息'].map('{:,.0f}'.format)
                st.dataframe(display_df, hide_index=True, use_container_width=True)
                
                st.markdown("### 📊 投資組合健康度")
                p1, p2 = st.columns(2)
                with p1:
                    fig_rating = px.pie(portfolio, names='Rating_Source', values='Weight', title='信評分佈')
                    st.plotly_chart(fig_rating, use_container_width=True)
                with p2:
                    if 'Issuer_Clean' in portfolio.columns: pie_col = 'Issuer_Clean'
                    else: pie_col = 'Name'
                    issuer_weights = portfolio.groupby(pie_col)['Weight'].sum().reset_index().sort_values('Weight', ascending=False)
                    if len(issuer_weights) > 5:
                        top5 = issuer_weights.head(5)
                        others = pd.DataFrame([{pie_col: 'Others', 'Weight': issuer_weights.iloc[5:]['Weight'].sum()}])
                        issuer_weights = pd.concat([top5, others])
                    fig_issuer = px.pie(issuer_weights, names=pie_col, values='Weight', title='發行機構分佈 (Smart Grouping)')
                    st.plotly_chart(fig_issuer, use_container_width=True)

            with c2:
                # 判斷是否顯示價差圖
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
                    fig_cf = px.bar(cf_df, x='Month', y='Amount', text_auto=',.0f', title=f"本金 ${investment_amt:,.0f} 之現金流模擬")
                    fig_cf.update_traces(marker_color='#2ecc71')
                    st.plotly_chart(fig_cf, use_container_width=True)
                    with st.expander("查看詳細配息日曆"):
                        st.dataframe(cf_detail_df, use_container_width=True)
                
                with my_tabs[2]:
                    st.caption(f"使用 **平均存續期間 ({avg_duration:.2f}年)** 進行利率敏感度分析 (基於原始資料)")
                    fig_risk = go.Figure()
                    text_positions = ['outside' if val < 0 else 'inside' for val in df_risk['資本損益']]
                    fig_risk.add_trace(go.Bar(
                        x=df_risk['情境'], y=df_risk['資本損益'], 
                        name='資本損益 (不含息)', 
                        marker_color=['#e74c3c' if x < 0 else '#2ecc71' for x in df_risk['資本損益']], 
                        text=df_risk['資本漲跌幅'], 
                        textposition=text_positions
                    ))
                    fig_risk.add_trace(go.Bar(x=df_risk['情境'], y=df_risk['利息收入'], name='利息收入 (預估一年)', marker_color='#3498db'))
                    fig_risk.add_trace(go.Scatter(x=df_risk['情境'], y=df_risk['總報酬'], name='總報酬 (含息)', mode='lines+markers+text', line=dict(color='gold', width=3), text=df_risk['總報酬漲跌幅'], textposition="top center"))
                    fig_risk.update_layout(barmode='relative', title="利率敏感度分析 (含漲跌幅 %)")
                    st.plotly_chart(fig_risk, use_container_width=True)

else:
    st.info("👆 請在上方選擇「公用報價檔」或「上傳新檔案」以開始分析。")

st.markdown("---")
st.markdown("""
<div style='background-color: #ffe6e6; padding: 10px; border-radius: 5px; color: #cc0000;'>
    <strong>⚠️ 投資風險警語 (Disclaimer)</strong><br>
    1. 本工具僅供投資試算與模擬使用，不代表任何形式之投資建議或獲利保證。<br>
    2. 債券價格、殖利率與配息金額均會隨市場波動，實際交易價格與條件請以銀行當下報價為準。<br>
    3. 投資人應自行評估風險承受能力，並詳閱公開說明書。外幣投資需自行承擔匯率風險。<br>
    4. 本系統之理論價格與價差分析僅為數學模型推估，非市場實際成交價格。<br>
    5. 本系統之風險試算採用您上傳之「存續期間」進行估算。
</div>
""", unsafe_allow_html=True)

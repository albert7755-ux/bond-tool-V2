import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import linprog, curve_fit
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# --- 1. 基礎設定 ---
st.set_page_config(page_title="債券策略大師 Pro (銀行實戰版)", layout="wide")

# 標題區
st.title("🛡️ 債券投資組合策略大師 Pro")
st.markdown("""
針對高資產客戶設計的策略模組：
1. **收益最大化**：追求最高配息。
2. **債券梯**：依據剩餘年期佈局，打造穩定現金流。
3. **槓鈴策略**：長短年期配置。
4. **相對價值**：找出「市價 < 理論價」的被低估債券。
5. **領息頻率組合**：自訂本金與領息頻率。
""")
st.divider()

# --- 2. 輔助函式 (維持不變，確保運算邏輯正確) ---
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
            return None, f"缺少必要欄位。請確認檔案包含：ISIN, 名稱, YTM, 剩餘年期。"

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
        
        if 'Original_Price' in df.columns:
            df['Valuation_Gap'] = df['Theoretical_Price'] - df['Original_Price']
        else:
            df['Original_Price'] = df['Theoretical_Price']
            df['Valuation_Gap'] = 0

        # 月份
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

# --- 3. 策略邏輯函數 (維持原樣) ---

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
                used_issuers.add(row['Name'])
                break
    if selected: return pd.DataFrame(selected)
    return pd.DataFrame()

def run_barbell(df, short_limit, long_limit, long_weight, allow_dup):
    short_pool = df[df['Years_Remaining'] <= short_limit].sort_values('YTM', ascending=False)
    long_pool = df[df['Years_Remaining'] >= long_limit].sort_values('YTM', ascending=False)
    selected, used_issuers = [], set()
    short_picks = []
    for idx, row in short_pool.iterrows():
        if len(short_picks) >= 2: break
        if allow_dup or (row['Name'] not in used_issuers):
            row = row.copy()
            row['Weight'] = (1 - long_weight) / 2 
            short_picks.append(row)
            used_issuers.add(row['Name'])
    long_picks = []
    for idx, row in long_pool.iterrows():
        if len(long_picks) >= 2: break
        if allow_dup or (row['Name'] not in used_issuers):
            row = row.copy()
            row['Weight'] = long_weight / 2
            long_picks.append(row)
            used_issuers.add(row['Name'])
    final_list = short_picks + long_picks
    if final_list: return pd.DataFrame(final_list)
    return pd.DataFrame()

def fit_yield_curve(x, a, b):
    return a + b * np.log(x)

def run_relative_value(df, allow_dup, top_n, min_dur, target_freqs):
    df_calc = df[df['Years_Remaining'] > 0.1].copy()
    if len(df_calc) < 5: return pd.DataFrame(), pd.DataFrame()
    
    pool = df_calc[df_calc['Years_Remaining'] >= min_dur]
    if target_freqs: pool = pool[pool['Frequency'].isin(target_freqs)]
    
    # 用價差排序
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
            
    try:
        popt, _ = curve_fit(fit_yield_curve, df_calc['Years_Remaining'], df_calc['YTM'])
        df_calc['Fair_YTM'] = fit_yield_curve(df_calc['Years_Remaining'], *popt)
    except:
        z = np.polyfit(df_calc['Years_Remaining'], df_calc['YTM'], 2)
        p = np.poly1d(z)
        df_calc['Fair_YTM'] = p(df_calc['Years_Remaining'])

    if selected: return pd.DataFrame(selected), df_calc
    return pd.DataFrame(), df_calc

def run_cash_flow_strategy(df, allow_dup, freq_type):
    selected = []
    used_issuers = set()
    if freq_type == "月月配 (12次/年)": target_months = [1, 2, 3, 4, 5, 6]
    elif freq_type == "雙月配 (6次/年)": target_months = [1, 3, 5]
    else: target_months = [1, 4]
    weight_per_bond = 1.0 / len(target_months)
    for m in target_months:
        pool = df[df['Pay_Month'] == m].sort_values('YTM', ascending=False)
        found = False
        for idx, row in pool.iterrows():
            if allow_dup or (row['Name'] not in used_issuers):
                bond = row.copy()
                bond['Weight'] = weight_per_bond
                bond['Cycle_Str'] = f"{m}月/{m+6}月"
                selected.append(bond)
                used_issuers.add(row['Name'])
                found = True
                break
    if selected: return pd.DataFrame(selected)
    return pd.DataFrame()

# --- 4. 主程式 UI (修正佈局) ---

# 【修改點 1】將檔案上傳區移到主畫面正中央，而不是側邊欄
st.subheader("📂 步驟 1: 請先上傳債券清單")
uploaded_file = st.file_uploader("支援銀行 Excel / CSV 格式", type=['xlsx', 'csv'])

if uploaded_file:
    df_raw, err = clean_data(uploaded_file)
    if err:
        st.error(f"錯誤: {err}")
    else:
        st.success(f"✅ 成功讀取 {len(df_raw)} 檔債券資料！請在左側選擇策略。")
        
        # --- 側邊欄：策略設定區 ---
        st.sidebar.header("🧠 步驟 2: 策略設定")
        
        # 黑名單
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

        # --- 策略參數與執行 ---
        if strategy == "收益最大化":
            t_dur = st.sidebar.slider("剩餘年期上限", 2.0, 30.0, 10.0)
            t_cred = rating_map[st.sidebar.select_slider("最低信評", list(rating_map.keys()), 'BBB')]
            max_w = st.sidebar.slider("單檔上限", 0.05, 0.5, 0.2)
            if st.sidebar.button("🚀 開始計算", type="primary"):
                portfolio = run_max_yield(df_clean, t_dur, t_cred, max_w)

        elif strategy == "債券梯":
            ladder_type = st.sidebar.selectbox("梯型結構", ["短梯 (1-5年)", "中梯 (3-7年)", "長梯 (5-15年)"])
            ladder_map = {"短梯 (1-5年)": [(1,2),(2,3),(3,4),(4,5)], "中梯 (3-7年)": [(3,4),(4,5),(5,6),(6,7)], "長梯 (5-15年)": [(5,7),(7,10),(10,12),(12,15)]}
            if st.sidebar.button("🚀 開始計算", type="primary"):
                portfolio = run_ladder(df_clean, ladder_map[ladder_type], allow_dup)

        elif strategy == "槓鈴策略":
            short_lim = st.sidebar.number_input("短債 < 年", 3.0)
            long_lim = st.sidebar.number_input("長債 > 年", 10.0)
            long_w = st.sidebar.slider("長債佔比", 0.1, 0.9, 0.5)
            if st.sidebar.button("🚀 開始計算", type="primary"):
                portfolio = run_barbell(df_clean, short_lim, long_lim, long_w, allow_dup)

        elif strategy == "相對價值":
            min_dur = st.sidebar.number_input("最低剩餘年期", 2.0)
            top_n = st.sidebar.slider("挑選幾檔", 3, 10, 5)
            target_rating = st.sidebar.multiselect("篩選信評 (推薦)", sorted(df_clean['Rating_Source'].unique()))
            available_freqs = sorted(df_clean['Frequency'].unique())
            target_freqs = st.sidebar.multiselect("篩選配息頻率", options=available_freqs, placeholder="全選")
            
            if st.sidebar.button("🚀 開始計算", type="primary"):
                df_t = df_clean[df_clean['Rating_Source'].isin(target_rating)] if target_rating else df_clean
                portfolio, df_with_alpha = run_relative_value(df_t, allow_dup, top_n, min_dur, target_freqs)

        elif strategy == "領息頻率組合":
            st.sidebar.caption("利用不同月份的半年配債券，構建現金流。")
            freq_type = st.sidebar.selectbox("目標領息頻率", ["月月配 (12次/年)", "雙月配 (6次/年)", "季季配 (4次/年)"])
            if df_clean['Is_Simulated_Month'].iloc[0]:
                st.sidebar.warning("⚠️ 警告：無法解析「到期日」，目前使用模擬月份。")
            if st.sidebar.button("🚀 開始計算", type="primary"):
                portfolio = run_cash_flow_strategy(df_clean, allow_dup, freq_type)

        # --- 結果顯示 ---
        if not portfolio.empty:
            st.divider()
            
            # 計算數據
            portfolio['Allocation %'] = (portfolio['Weight'] * 100).round(1)
            price_col = 'Original_Price' if 'Original_Price' in portfolio.columns else 'Theoretical_Price'
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
            avg_years = (portfolio['Years_Remaining'] * portfolio['Weight']).sum()
            
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("預期年化殖利率", f"{avg_ytm:.2f}%")
            k2.metric("平均剩餘年期", f"{avg_years:.2f} 年")
            k3.metric("預估年領總息", f"${total_coupon:,.0f}")
            k4.metric("平均買入價格", f"${avg_price:.2f}")

            c1, c2 = st.columns([5, 5])
            with c1:
                st.subheader("📋 建議清單 (含價差分析)")
                cols = ['Name', 'YTM', 'Years_Remaining', 'Allocation %', 'Annual_Coupon_Amt']
                if 'Theoretical_Price' in portfolio.columns: cols.insert(2, 'Theoretical_Price')
                if 'Original_Price' in portfolio.columns: cols.insert(3, 'Original_Price')
                if 'Valuation_Gap' in portfolio.columns: cols.insert(4, 'Valuation_Gap')
                if 'Frequency' in portfolio.columns: cols.append('Frequency')
                if 'Cycle_Str' in portfolio.columns: cols.insert(1, 'Cycle_Str')
                
                rename_dict = {
                    'Original_Price': '銀行報價 (Offer)',
                    'Theoretical_Price': '理論價格',
                    'Valuation_Gap': '潛在價差 (Alpha)',
                    'Years_Remaining': '剩餘年期',
                    'Annual_Coupon_Amt': '預估年息'
                }
                display_df = portfolio[cols].rename(columns=rename_dict).copy()
                
                # 格式化
                for c in ['銀行報價 (Offer)', '理論價格', '潛在價差 (Alpha)', '剩餘年期']:
                    if c in display_df.columns: display_df[c] = display_df[c].map('{:.2f}'.format)
                if '預估年息' in display_df.columns: display_df['預估年息'] = display_df['預估年息'].map('{:,.0f}'.format)
                
                st.dataframe(display_df, hide_index=True, use_container_width=True)

            with c2:
                tab1, tab2 = st.tabs(["📊 策略分析", "💰 現金流試算"])
                
                with tab1:
                    if strategy == "相對價值" and not df_with_alpha.empty:
                        st.subheader("相對價值回歸分析")
                        base_data = df_with_alpha
                        x_range = np.linspace(base_data['Years_Remaining'].min(), base_data['Years_Remaining'].max(), 100)
                        try:
                            popt, _ = curve_fit(fit_yield_curve, base_data['Years_Remaining'], base_data['YTM'])
                            y_fair = fit_yield_curve(x_range, *popt)
                        except:
                            z = np.polyfit(base_data['Years_Remaining'], base_data['YTM'], 2)
                            p = np.poly1d(z)
                            y_fair = p(x_range)
                        
                        fig_rv = go.Figure()
                        fig_rv.add_trace(go.Scatter(x=base_data['Years_Remaining'], y=base_data['YTM'], mode='markers', name='市場', marker=dict(color='lightgrey', size=6), hovertext=base_data['Name']))
                        fig_rv.add_trace(go.Scatter(x=x_range, y=y_fair, mode='lines', name='合理價值', line=dict(dash='dash', color='blue')))
                        fig_rv.add_trace(go.Scatter(x=portfolio['Years_Remaining'], y=portfolio['YTM'], mode='markers', name='低估買入', marker=dict(color='red', size=15, symbol='star'), hovertext=portfolio['Name']))
                        fig_rv.update_layout(xaxis_title="剩餘年期 (Years)", yaxis_title="殖利率 (YTM)")
                        st.plotly_chart(fig_rv, use_container_width=True)
                        
                    elif strategy == "領息頻率組合":
                         st.info("👈 請切換至「現金流試算」分頁查看詳細圖表")
                    else:
                        st.subheader("風險/收益分佈圖")
                        df_raw['Type'] = '未選入'
                        portfolio['Type'] = '建議買入'
                        if excluded_issuers: df_raw.loc[df_raw['Name'].isin(excluded_issuers), 'Type'] = '已剔除'
                        all_plot = pd.concat([df_raw[~df_raw['ISIN'].isin(portfolio['ISIN'])], portfolio])
                        color_map = {'未選入': '#e0e0e0', '建議買入': '#ef553b', '已剔除': 'rgba(0,0,0,0.1)'}
                        fig = px.scatter(
                            all_plot, x='Years_Remaining', y='YTM', color='Type', 
                            color_discrete_map=color_map,
                            size=all_plot['Type'].map({'未選入': 5, '建議買入': 15, '已剔除': 3}),
                            hover_data=['Name'],
                            title=f"{strategy} 策略分佈",
                            labels={'Years_Remaining': '剩餘年期 (Years)'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

                with tab2:
                    st.subheader("預估每月入帳金額 (稅前)")
                    months = list(range(1, 13))
                    cash_flow = [0] * 12
                    for idx, row in portfolio.iterrows():
                        freq_val = row.get('Frequency', '半年配')
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
                        else:
                            per_pay = coupon_amt / 2
                            cash_flow[m_idx] += per_pay
                            cash_flow[(m_idx + 6) % 12] += per_pay
                    
                    cf_df = pd.DataFrame({'Month': [f"{i}月" for i in months], 'Amount': cash_flow})
                    fig_cf = px.bar(cf_df, x='Month', y='Amount', text_auto=',.0f', title=f"本金 ${investment_amt:,.0f} 之現金流模擬")
                    fig_cf.update_traces(marker_color='#2ecc71')
                    fig_cf.update_layout(yaxis_title="金額 (元)")
                    st.plotly_chart(fig_cf, use_container_width=True)

        elif uploaded_file and st.session_state.get('last_run'):
            st.warning("⚠️ 找不到符合條件的債券。")

else:
    # 這是為了讓還沒上傳時，畫面不會太乾淨，給一些指引
    st.info("👆 請在上方上傳您的債券清單 Excel 檔以開始分析。")

# 【修改點 2】新增合規警語 (Footer)
st.markdown("---")
st.markdown("""
<div style='background-color: #ffe6e6; padding: 10px; border-radius: 5px; color: #cc0000;'>
    <strong>⚠️ 投資風險警語 (Disclaimer)</strong><br>
    1. 本工具僅供投資試算與模擬使用，不代表任何形式之投資建議或獲利保證。<br>
    2. 債券價格、殖利率與配息金額均會隨市場波動，實際交易價格與條件請以銀行當下報價為準。<br>
    3. 投資人應自行評估風險承受能力，並詳閱公開說明書。外幣投資需自行承擔匯率風險。<br>
    4. 本系統之理論價格與價差分析僅為數學模型推估，非市場實際成交價格。
</div>
""", unsafe_allow_html=True)

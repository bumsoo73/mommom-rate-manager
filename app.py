import streamlit as st
import pandas as pd
from datetime import timedelta, date, datetime
import calendar
import io
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Page Basic Settings ---
st.set_page_config(layout="wide", page_title="호텔 상품 관리 시스템 Final (Google Sheets)")

# --- Custom Styles ---
st.markdown("""
    <style>
    /* Button Styles */
    .stButton>button[kind="secondary"] {
        color: #e65100 !important; border: none !important; background: transparent !important; box-shadow: none !important;
    }
    .stButton>button[kind="secondary"]:hover {
        color: #ef6c00 !important; background-color: #fff3e0 !important;
    }
    .stButton>button[kind="primary"] {
        background-color: #ef6c00 !important; border-color: #ef6c00 !important; color: white !important;
    }
    
    /* Calendar & Table Styles */
    .calendar-table { width: 100%; border-collapse: collapse; }
    .calendar-table th { background-color: #fff3e0; padding: 10px; text-align: center; border: 1px solid #ddd; color: #e65100; }
    .calendar-table td { vertical-align: top; height: 120px; border: 1px solid #ddd; padding: 5px; width: 14%; }
    .day-number { font-weight: bold; margin-bottom: 5px; display: block; color: #555; }
    .prod-item { font-size: 0.8em; background-color: #fff8e1; margin-bottom: 2px; padding: 2px 4px; border-radius: 3px; color: #bf360c; border: 1px solid #ffe0b2; }
    
    /* Tags */
    .price-tag { font-weight: bold; color: #ef6c00; }
    .stock-tag { font-weight: bold; color: #1565c0; background-color: #e3f2fd; padding: 1px 4px; border-radius: 4px; font-size: 0.9em; }
    .stock-zero { font-weight: bold; color: #b71c1c; background-color: #ffcdd2; border: 1px solid #ef9a9a; padding: 1px 4px; border-radius: 4px; font-size: 0.9em; }
    .other-month { background-color: #f9f9f9; color: #ccc; }
    
    /* UI Fixes */
    div[data-testid="column"] button[kind="secondary"] { border: 0px solid transparent !important; background: transparent !important; }
    [data-testid="stCheckbox"] { margin-right: 0px; padding-right: 0px; }
    </style>
    """, unsafe_allow_html=True)

# --- [CORE] Google Sheets Connection ---
@st.cache_resource
def connect_to_gsheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # secrets.toml 파일이 없으면 여기서 에러가 납니다.
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

def get_sheet_data():
    client = connect_to_gsheet()
    try:
        sh = client.open("mommom-rate-manager")
    except gspread.SpreadsheetNotFound:
        st.error("🚨 'mommom-rate-manager' 구글 시트를 찾을 수 없습니다. (공유 설정을 확인하세요)")
        return [], [], pd.DataFrame()

    try:
        ws_hotels = sh.worksheet("hotels")
        hotels_data = ws_hotels.get_all_values()
        hotels = [row[0] for row in hotels_data[1:]] if len(hotels_data) > 1 else ["쏠비치 삼척", "소노벨 천안"]
    except: hotels = ["쏠비치 삼척", "소노벨 천안"]

    try:
        ws_products = sh.worksheet("products")
        products_data = ws_products.get_all_records()
        products = products_data if products_data else [
            {'hotel': '쏠비치 삼척', 'name': '[3인] 패밀리 스탠다드'},
            {'hotel': '쏠비치 삼척', 'name': '[4인] 스위트 오션'}
        ]
    except: products = []

    try:
        ws_main = sh.worksheet("main_data")
        main_data = ws_main.get_all_records()
        main_df = pd.DataFrame(main_data)
        if not main_df.empty:
            main_df['날짜'] = pd.to_datetime(main_df['날짜']).dt.date
        else:
            main_df = pd.DataFrame(columns=['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])
    except:
        main_df = pd.DataFrame(columns=['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])

    return hotels, products, main_df

def save_to_gsheet(sheet_type, data):
    client = connect_to_gsheet()
    sh = client.open("mommom-rate-manager")
    
    if sheet_type == 'hotels':
        ws = sh.worksheet("hotels")
        ws.clear()
        ws.update([["숙소명"]] + [[h] for h in data])
        
    elif sheet_type == 'products':
        ws = sh.worksheet("products")
        ws.clear()
        if data:
            headers = list(data[0].keys())
            values = [list(d.values()) for d in data]
            ws.update([headers] + values)
        else:
            ws.update([["hotel", "name"]])
            
    elif sheet_type == 'main_data':
        ws = sh.worksheet("main_data")
        ws.clear()
        if not data.empty:
            save_df = data.copy()
            save_df['날짜'] = save_df['날짜'].astype(str)
            update_data = [save_df.columns.values.tolist()] + save_df.values.tolist()
            ws.update(update_data)
        else:
            ws.update([['날짜', '숙소명', '상품명', '요금', '재고', '판매상태']])

# --- Initialize ---
if 'data_loaded' not in st.session_state:
    with st.spinner("구글 시트 로딩 중..."):
        hotels, products, main_df = get_sheet_data()
        st.session_state.hotels = hotels
        st.session_state.products = products
        st.session_state.main_df = main_df
        st.session_state.data_loaded = True

if 'selected_dates_buffer' not in st.session_state:
    st.session_state.selected_dates_buffer = []
if 'cal_year' not in st.session_state:
    st.session_state.cal_year = date.today().year
if 'cal_month' not in st.session_state:
    st.session_state.cal_month = date.today().month
if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False
if 'hotel_to_delete' not in st.session_state:
    st.session_state.hotel_to_delete = None

# --- Helpers ---
def format_date_kr(d):
    if isinstance(d, str):
        try: d = pd.to_datetime(d).date()
        except: return d
    elif isinstance(d, datetime): d = d.date()
    weekdays = ["월", "화", "수", "목", "금", "토", "일"]
    return f"{d.year}-{d.month:02d}-{d.day:02d} ({weekdays[d.weekday()]})"

def generate_dates(start_date, end_date, weekdays):
    dates = []
    curr = start_date
    while curr <= end_date:
        if curr.weekday() in weekdays: dates.append(curr)
        curr += timedelta(days=1)
    return dates

def change_month(amount):
    st.session_state.cal_month += amount
    if st.session_state.cal_month > 12:
        st.session_state.cal_month = 1
        st.session_state.cal_year += 1
    elif st.session_state.cal_month < 1:
        st.session_state.cal_month = 12
        st.session_state.cal_year -= 1

# Product Logic
def move_product(current_hotel, index, direction):
    all = st.session_state.products
    curr_prods = [p for p in all if p['hotel'] == current_hotel]
    others = [p for p in all if p['hotel'] != current_hotel]
    
    if direction == -1 and index > 0:
        curr_prods[index], curr_prods[index-1] = curr_prods[index-1], curr_prods[index]
    elif direction == 1 and index < len(curr_prods) - 1:
        curr_prods[index], curr_prods[index+1] = curr_prods[index+1], curr_prods[index]
    
    st.session_state.products = others + curr_prods
    save_to_gsheet('products', st.session_state.products)

def delete_product(current_hotel, index):
    all = st.session_state.products
    curr_prods = [p for p in all if p['hotel'] == current_hotel]
    others = [p for p in all if p['hotel'] != current_hotel]
    del curr_prods[index]
    st.session_state.products = others + curr_prods
    save_to_gsheet('products', st.session_state.products)

# ==========================================
# UI Layout
# ==========================================
with st.sidebar:
    st.title("🏢 숙소 선택")
    search_query = st.text_input("🔍 숙소 검색", placeholder="숙소명을 입력하세요")
    
    if st.session_state.hotels:
        filtered_hotels = [h for h in st.session_state.hotels if search_query in h] if search_query else st.session_state.hotels
        if filtered_hotels:
            current_hotel = st.selectbox("작업할 숙소 선택", filtered_hotels, index=0)
        else:
            st.warning("검색된 숙소가 없습니다.")
            current_hotel = None
    else:
        current_hotel = None
        st.warning("등록된 숙소가 없습니다.")

    st.markdown("---")
    with st.expander("⚙️ 숙소 리스트 관리"):
        tab_add, tab_del = st.tabs(["추가", "삭제"])
        with tab_add:
            with st.form("add_hotel_form", clear_on_submit=True):
                new_h = st.text_input("새 숙소명 입력")
                if st.form_submit_button("➕ 추가하기"):
                    if new_h and new_h not in st.session_state.hotels:
                        st.session_state.hotels.append(new_h)
                        save_to_gsheet('hotels', st.session_state.hotels)
                        st.success(f"'{new_h}' 추가됨")
                        st.rerun()
        with tab_del:
            if st.button("🗑 선택한 숙소 삭제"):
                st.session_state.confirm_delete = True
                st.session_state.hotel_to_delete = current_hotel # 현재 선택된 숙소 삭제 시도
            
            if st.session_state.confirm_delete:
                st.error(f"정말 '{st.session_state.hotel_to_delete}'를 삭제하시겠습니까?")
                c_y, c_n = st.columns(2)
                if c_y.button("✅ 예"):
                    target = st.session_state.hotel_to_delete
                    st.session_state.hotels.remove(target)
                    st.session_state.products = [p for p in st.session_state.products if p['hotel'] != target]
                    st.session_state.main_df = st.session_state.main_df[st.session_state.main_df['숙소명'] != target]
                    
                    save_to_gsheet('hotels', st.session_state.hotels)
                    save_to_gsheet('products', st.session_state.products)
                    save_to_gsheet('main_data', st.session_state.main_df)
                    
                    st.session_state.confirm_delete = False
                    st.success("삭제되었습니다.")
                    st.rerun()
                if c_n.button("❌ 아니오"):
                    st.session_state.confirm_delete = False
                    st.rerun()

if current_hotel:
    st.header(f"🏨 {current_hotel} 관리")
    tab_prod, tab_work, tab_excel = st.tabs(["1. 📦 상품 세팅", "2. 📅 가격/재고 등록 & 확인", "3. 📤 엑셀 추출"])

    # TAB 1: Product
    with tab_prod:
        c1, c2 = st.columns([1, 1.5], gap="large")
        with c1:
            st.subheader("상품 등록")
            with st.form("add_prod", clear_on_submit=True):
                new_p = st.text_input("상품명 (객실타입)")
                if st.form_submit_button("상품 추가"):
                    my_prods = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
                    if new_p and new_p not in my_prods:
                        st.session_state.products.append({'hotel': current_hotel, 'name': new_p})
                        save_to_gsheet('products', st.session_state.products)
                        st.success("추가 완료")
                        st.rerun()
                    elif new_p in my_prods: st.warning("이미 존재함")
        with c2:
            st.subheader("등록된 상품 순서 관리")
            st.caption("⬆️ ⬇️ 버튼을 눌러 순서를 변경하세요.")
            curr_list = [p for p in st.session_state.products if p['hotel'] == current_hotel]
            if curr_list:
                for i, prod in enumerate(curr_list):
                    c_a, c_b, c_c, c_d = st.columns([0.5, 0.5, 4, 0.5])
                    with c_a:
                        if i > 0: 
                            if st.button("⬆️", key=f"u_{i}"): move_product(current_hotel, i, -1); st.rerun()
                    with c_b:
                        if i < len(curr_list)-1:
                            if st.button("⬇️", key=f"d_{i}"): move_product(current_hotel, i, 1); st.rerun()
                    with c_c: st.markdown(f"<div style='padding-top:5px;'><b>{prod['name']}</b></div>", unsafe_allow_html=True)
                    with c_d:
                        if st.button("🗑️", key=f"del_{i}"): delete_product(current_hotel, i); st.rerun()
                    st.divider()
            else: st.info("상품 없음")

    # TAB 2: Work
    with tab_work:
        with st.expander("⚡️ 가격/재고 일괄 입력 열기", expanded=True):
            my_products = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
            if not my_products: st.warning("상품 먼저 등록 필요")
            else:
                st.markdown("#### 1. 날짜 및 요일 선택")
                c_d1, c_d2 = st.columns([1, 2])
                with c_d1:
                    d_range = st.date_input("기간", [], help="시작/종료일 선택")
                    if len(d_range) == 2: st.success(f"{format_date_kr(d_range[0])} ~ {format_date_kr(d_range[1])}")
                with c_d2:
                    st.write("요일 필터")
                    days = ["일","월","화","수","목","금","토"]
                    sel_days = []
                    cols = st.columns([1]*7 + [10])
                    for i, d in enumerate(days):
                        if cols[i].checkbox(d, True, key=f"dw_{i}"): sel_days.append(6 if i==0 else i-1) # Py: Mon=0, Sun=6
                
                if st.button("⬇️ 기간 추가", type="secondary"):
                    if len(d_range)!=2: st.error("기간 선택 필요")
                    elif not sel_days: st.error("요일 선택 필요")
                    else:
                        new_ds = generate_dates(d_range[0], d_range[1], sel_days)
                        if not new_ds: st.warning("해당 기간에 선택한 요일이 없습니다.")
                        else:
                            buf = set(st.session_state.selected_dates_buffer)
                            for d in new_ds: buf.add(format_date_kr(d))
                            st.session_state.selected_dates_buffer = sorted(list(buf))
                            st.rerun()
                
                if st.session_state.selected_dates_buffer:
                    st.markdown("---")
                    st.markdown("##### ✅ 적용 대상 날짜")
                    upd = st.multiselect("선택된 날짜들", st.session_state.selected_dates_buffer, st.session_state.selected_dates_buffer, label_visibility="collapsed")
                    if len(upd) != len(st.session_state.selected_dates_buffer):
                        st.session_state.selected_dates_buffer = upd
                        st.rerun()

                st.markdown("---")
                st.markdown("#### 2. 상품별 설정")
                sel_prods = st.multiselect("작업할 상품", my_products, my_products)
                input_map = {}
                for p in sel_prods:
                    st.markdown(f"**🔹 {p}**")
                    pc1, pc2, pc3 = st.columns(3)
                    pr = pc1.number_input("요금", key=f"p_{p}", step=1000, value=None)
                    stk = pc2.number_input("재고", key=f"s_{p}", value=5)
                    sts = pc3.selectbox("상태", ["Y","N"], key=f"st_{p}")
                    input_map[p] = {'p':pr, 's':stk, 'st':sts}
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 데이터 생성하기 (최종 저장)", type="primary", use_container_width=True):
                    if not st.session_state.selected_dates_buffer: st.error("날짜를 추가해주세요.")
                    elif not sel_prods: st.error("상품을 선택해주세요.")
                    else:
                        missing = False
                        for p, v in input_map.items():
                            if v['p'] is None: missing=True; break
                        if missing: st.error("요금을 입력해주세요.")
                        else:
                            final_ds = []
                            for s in st.session_state.selected_dates_buffer:
                                # YYYY-MM-DD (요일) -> YYYY-MM-DD
                                d_str = s.split(" ")[0]
                                try: final_ds.append(datetime.strptime(d_str, "%Y-%m-%d").date())
                                except: pass
                            
                            new_rows = []
                            for d in final_ds:
                                for p, v in input_map.items():
                                    new_rows.append({
                                        '날짜': d, '숙소명': current_hotel, '상품명': p,
                                        '요금': v['p'], '재고': v['s'], '판매상태': v['st']
                                    })
                            
                            st.session_state.main_df = pd.concat([st.session_state.main_df, pd.DataFrame(new_rows)], ignore_index=True)
                            st.session_state.main_df['날짜'] = pd.to_datetime(st.session_state.main_df['날짜']).dt.date
                            st.session_state.main_df.drop_duplicates(subset=['날짜','숙소명','상품명'], keep='last', inplace=True)
                            st.session_state.main_df.sort_values(['날짜','상품명'], inplace=True)
                            
                            save_to_gsheet('main_data', st.session_state.main_df)
                            st.session_state.selected_dates_buffer = []
                            st.success("저장 완료!")
                            st.rerun()

        st.divider()
        st.markdown("### 📊 데이터 확인 및 수정")
        
        curr_order = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
        hotel_df = st.session_state.main_df[st.session_state.main_df['숙소명'] == current_hotel].copy()
        if not hotel_df.empty and curr_order:
            hotel_df['상품명'] = pd.Categorical(hotel_df['상품명'], curr_order, ordered=True)
            hotel_df = hotel_df.sort_values(['날짜', '상품명'])
            
        view = st.radio("보기 방식", ["📋 리스트 표보기", "🗓️ 월별 요금 캘린더", "📅 월별 재고 캘린더"], horizontal=True)
        
        if "리스트" in view:
            if hotel_df.empty: st.info("데이터 없음")
            else:
                edited = st.data_editor(
                    hotel_df[['날짜','상품명','요금','재고','판매상태']],
                    column_config={
                        "날짜": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
                        "상품명": st.column_config.TextColumn(disabled=True),
                        "요금": st.column_config.NumberColumn(format="%d"),
                    },
                    use_container_width=True, hide_index=True
                )
                if not edited.equals(hotel_df[['날짜','상품명','요금','재고','판매상태']]):
                    # [Fix] st.session_state로 수정 완료
                    st.session_state.main_df.loc[edited.index, ['날짜','요금','재고','판매상태']] = edited[['날짜','요금','재고','판매상태']]
                    save_to_gsheet('main_data', st.session_state.main_df)
                    st.toast("수정 저장됨")
        else:
            is_stock = "재고" in view
            c_p, c_t, c_n, _ = st.columns([1, 4, 1, 6])
            if c_p.button("⬅️"): change_month(-1); st.rerun()
            if c_n.button("➡️"): change_month(1); st.rerun()
            curr_y, curr_m = st.session_state.cal_year, st.session_state.cal_month
            c_t.markdown(f"#### {curr_y}년 {curr_m}월")
            
            mask = (pd.to_datetime(hotel_df['날짜']).dt.year == curr_y) & (pd.to_datetime(hotel_df['날짜']).dt.month == curr_m)
            m_data = hotel_df[mask]
            
            calendar.setfirstweekday(calendar.SUNDAY)
            cal = calendar.monthcalendar(curr_y, curr_m)
            html = "<table class='calendar-table'><thead><tr>" + "".join([f"<th>{d}</th>" for d in ["일","월","화","수","목","금","토"]]) + "</tr></thead><tbody>"
            
            for week in cal:
                html += "<tr>"
                for d in week:
                    if d==0: html += "<td class='other-month'></td>"
                    else:
                        d_obj = date(curr_y, curr_m, d)
                        recs = m_data[pd.to_datetime(m_data['날짜']).dt.date == d_obj]
                        if not recs.empty and curr_order:
                            recs['상품명'] = pd.Categorical(recs['상품명'], curr_order, ordered=True)
                            recs = recs.sort_values('상품명')
                        
                        cell = f"<span class='day-number'>{d}</span>"
                        for _, r in recs.iterrows():
                            nm = r['상품명']
                            if is_stock:
                                q = r['재고']
                                val, cls = (f"{q}개", "stock-tag") if q>0 else ("품절", "stock-zero")
                            else:
                                val, cls = f"{r['요금']:,}", "price-tag"
                            cell += f"<div class='prod-item'>{nm}<br><span class='{cls}'>{val}</span></div>"
                        html += f"<td>{cell}</td>"
                html += "</tr>"
            html += "</tbody></table>"
            st.markdown(html, unsafe_allow_html=True)

    # TAB 3: Excel
    with tab_excel:
        st.subheader("엑셀 다운로드")
        hotel_df = st.session_state.main_df[st.session_state.main_df['숙소명'] == current_hotel].copy()
        if hotel_df.empty: st.warning("데이터 없음")
        else:
            if curr_order:
                hotel_df['상품명'] = pd.Categorical(hotel_df['상품명'], curr_order, ordered=True)
                hotel_df = hotel_df.sort_values(['날짜','상품명'])
            
            ex_data = []
            for _, r in hotel_df.iterrows():
                row = [""]*13
                try: row[0] = format_date_kr(r['날짜'])
                except: row[0] = str(r['날짜'])
                row[1] = r['상품명']; row[6] = r['요금']; row[8] = r['재고']; row[12] = r['판매상태']
                ex_data.append(row)
            
            df_ex = pd.DataFrame(ex_data, columns=["날짜(A)", "상품명(B)", "C", "D", "E", "F", "요금(G)", "H", "재고(I)", "J", "K", "L", "판매상태(M)"])
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: df_ex.to_excel(w, index=False, sheet_name='Sheet1')
            out.seek(0)
            st.download_button("📥 엑셀 다운로드", out, f"[{current_hotel}]_{date.today()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

else:
    st.info("왼쪽에서 숙소를 선택하세요.")

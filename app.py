import streamlit as st
import pandas as pd
from datetime import timedelta, date, datetime
import calendar
import io
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Page Basic Settings ---
st.set_page_config(layout="wide", page_title="맘맘 요금재고 관리툴")

# --- Custom Styles ---
st.markdown("""
    <style>
    /* Secondary Button (Transparent) */
    .stButton>button[kind="secondary"] {
        color: #e65100 !important; 
        border: none !important; 
        background-color: transparent !important; 
        box-shadow: none !important;
    }
    .stButton>button[kind="secondary"]:hover {
        color: #ef6c00 !important;
        background-color: #fff3e0 !important;
    }
    
    /* Primary Button (Orange - Data Save) */
    .stButton>button[kind="primary"] {
        background-color: #ff6d00 !important; 
        border-color: #ff6d00 !important;
        color: white !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #e65100 !important;
        border-color: #e65100 !important;
    }

    /* Black Button Style (For Delete & Add Period) */
    .black-btn > button {
        background-color: #212121 !important;
        border-color: #212121 !important;
        color: white !important;
        width: 100%;
    }
    .black-btn > button:hover {
        background-color: #424242 !important;
        border-color: #424242 !important;
    }

    /* Calendar & Table */
    .calendar-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .calendar-table th { background-color: #fff3e0; padding: 8px; text-align: center; border: 1px solid #ddd; color: #e65100; }
    .calendar-table td { vertical-align: top; height: 100px; border: 1px solid #ddd; padding: 5px; width: 14%; }
    .day-number { font-weight: bold; margin-bottom: 5px; display: block; color: #555; }
    .prod-item { font-size: 0.75em; background-color: #fff8e1; margin-bottom: 2px; padding: 2px 4px; border-radius: 3px; color: #bf360c; border: 1px solid #ffe0b2; }
    
    /* Tags */
    .price-tag { font-weight: bold; color: #ef6c00; }
    .stock-tag { font-weight: bold; color: #1565c0; background-color: #e3f2fd; padding: 1px 4px; border-radius: 4px; font-size: 0.9em; }
    .stock-zero { font-weight: bold; color: #b71c1c; background-color: #ffcdd2; border: 1px solid #ef9a9a; padding: 1px 4px; border-radius: 4px; font-size: 0.9em; }
    .stop-sales { font-weight: bold; color: white; background-color: #424242; padding: 1px 4px; border-radius: 4px; font-size: 0.85em; margin-right: 3px; }
    .other-month { background-color: #f9f9f9; color: #ccc; }
    
    /* UI Adjustments */
    div[data-testid="column"] button[kind="secondary"] { border: 0px solid transparent !important; background: transparent !important; }
    [data-testid="stCheckbox"] { margin-right: 0px; padding-right: 0px; }
    
    /* Selected Date Box */
    .selected-date-box {
        background-color: #e3f2fd;
        padding: 10px;
        border-radius: 5px;
        border: 1px solid #90caf9;
        color: #1565c0;
        font-weight: bold;
        text-align: center;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- [CORE] Google Sheets Connection ---
@st.cache_resource
def connect_to_gsheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

# --- Metadata Loader ---
def get_metadata():
    client = connect_to_gsheet()
    try: sh = client.open("Mammam_DB")
    except: return [], []

    try:
        ws_h = sh.worksheet("hotels")
        h_data = ws_h.get_all_values()
        hotels = [row[0] for row in h_data[1:]] if len(h_data) > 1 else []
    except: hotels = []

    try:
        ws_p = sh.worksheet("products")
        p_data = ws_p.get_all_records()
        products = p_data if p_data else []
    except: products = []

    return hotels, products

# --- Hotel Data Loader ---
def get_hotel_data(hotel_name):
    client = connect_to_gsheet()
    sh = client.open("Mammam_DB")
    sheet_name = f"DB_{hotel_name}"
    
    try:
        ws = sh.worksheet(sheet_name)
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if not df.empty:
            df['날짜'] = pd.to_datetime(df['날짜']).dt.date
        else:
            df = pd.DataFrame(columns=['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=10)
        ws.append_row(['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])
        df = pd.DataFrame(columns=['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])
    return df

# --- Save Logic ---
def save_metadata(type, data):
    client = connect_to_gsheet()
    sh = client.open("Mammam_DB")
    
    if type == 'hotels':
        try: ws = sh.worksheet("hotels")
        except: ws = sh.add_worksheet("hotels", 100, 1)
        ws.clear()
        ws.update([["숙소명"]] + [[h] for h in data])
        
    elif type == 'products':
        try: ws = sh.worksheet("products")
        except: ws = sh.add_worksheet("products", 100, 4)
        ws.clear()
        if data:
            for item in data:
                if 'code' not in item: item['code'] = ""
            headers = list(data[0].keys())
            values = [list(d.values()) for d in data]
            ws.update([headers] + values)
        else:
            ws.update([["hotel", "name", "code"]])

def save_hotel_data(hotel_name, df):
    client = connect_to_gsheet()
    sh = client.open("Mammam_DB")
    sheet_name = f"DB_{hotel_name}"
    
    try: ws = sh.worksheet(sheet_name)
    except: ws = sh.add_worksheet(sheet_name, 1000, 10)
    
    ws.clear()
    if not df.empty:
        save_df = df.copy()
        save_df['날짜'] = save_df['날짜'].astype(str)
        update_data = [save_df.columns.values.tolist()] + save_df.values.tolist()
        ws.update(update_data)
    else:
        ws.update([['날짜', '숙소명', '상품명', '요금', '재고', '판매상태']])

# --- Helpers ---
def format_date_kr(d):
    if isinstance(d, str):
        try: d = pd.to_datetime(d).date()
        except: return d
    elif isinstance(d, datetime): d = d.date()
    weekdays = ["월", "화", "수", "목", "금", "토", "일"]
    return f"{d.year}-{d.month:02d}-{d.day:02d} ({weekdays[d.weekday()]})"

def generate_dates(start, end, weekdays):
    dates = []
    curr = start
    while curr <= end:
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

def move_product(hotel, idx, direct):
    all_p = st.session_state.products
    curr = [p for p in all_p if p['hotel'] == hotel]
    others = [p for p in all_p if p['hotel'] != hotel]
    if direct == -1 and idx > 0: curr[idx], curr[idx-1] = curr[idx-1], curr[idx]
    elif direct == 1 and idx < len(curr)-1: curr[idx], curr[idx+1] = curr[idx+1], curr[idx]
    st.session_state.products = others + curr
    save_metadata('products', st.session_state.products)

def delete_product_item(hotel, idx):
    all_p = st.session_state.products
    curr = [p for p in all_p if p['hotel'] == hotel]
    others = [p for p in all_p if p['hotel'] != hotel]
    del curr[idx]
    st.session_state.products = others + curr
    save_metadata('products', st.session_state.products)

def update_download_log():
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if 'download_logs' not in st.session_state:
        st.session_state.download_logs = []
    st.session_state.download_logs.insert(0, now_str)

# --- Initialization ---
if 'init' not in st.session_state:
    with st.spinner("데이터 로딩 중..."):
        h, p = get_metadata()
        st.session_state.hotels = h if h else []
        st.session_state.products = p
        st.session_state.selected_dates_buffer = []
        st.session_state.cal_year = date.today().year
        st.session_state.cal_month = date.today().month
        st.session_state.confirm_delete_req = False
        st.session_state.input_reset_key = 0
        st.session_state.download_logs = []
        st.session_state.init = True

# ==========================================
# Sidebar
# ==========================================
with st.sidebar:
    st.title("맘맘 요금재고 관리툴") 
    
    st.markdown("### 🔍 숙소 검색")
    search_q = st.text_input("search_hotel", placeholder="검색어 입력 후 엔터", label_visibility="collapsed")
    st.caption("검색어 입력 후 아래 목록에서 선택")
    
    # [요청] 최신순 정렬
    sorted_hotels = st.session_state.hotels[::-1]
    filtered_hotels = [h for h in sorted_hotels if search_q in h] if search_q else sorted_hotels
    
    current_hotel = None
    if filtered_hotels:
        current_hotel = st.selectbox("숙소 선택", filtered_hotels)
        if 'last_hotel' not in st.session_state or st.session_state.last_hotel != current_hotel:
            with st.spinner(f"'{current_hotel}' 데이터 로딩 중..."):
                st.session_state.main_df = get_hotel_data(current_hotel)
                st.session_state.last_hotel = current_hotel
    else:
        st.warning("등록된 숙소가 없습니다.")

    st.markdown("---")
    
    # [요청] 토글 디폴트 펼침
    with st.expander("⚙️ 숙소 리스트 관리", expanded=True):
        t1, t2 = st.tabs(["추가", "삭제"])
        with t1:
            with st.form("add_h"):
                new_h = st.text_input("새 숙소명")
                if st.form_submit_button("추가"):
                    if new_h and new_h not in st.session_state.hotels:
                        st.session_state.hotels.append(new_h)
                        save_metadata('hotels', st.session_state.hotels)
                        get_hotel_data(new_h)
                        st.success(f"'{new_h}' 추가 완료")
                        time.sleep(1)
                        st.rerun()
        with t2:
            # [요청] 삭제 UI 개선: 버튼 위에 숙소명 표시
            if current_hotel:
                st.caption(f"현재 선택된 숙소: **{current_hotel}**")
                
                # [요청] 검정색 버튼 (Custom Class)
                c_del_btn = st.columns([1])[0]
                with c_del_btn:
                    # Streamlit button style hack using class mapping
                    # But easiest is to use a container with specific class or just native primary if acceptable.
                    # User wanted Black button. We defined .black-btn in CSS.
                    # Since Streamlit doesn't allow class assignment to button directly easily,
                    # We will use Primary button here for consistency as requested in previous turn for 'Period Add'
                    # But actually user said 'Current Hotel Delete' should be black.
                    # We will wrap it.
                    
                    if st.button("현재 숙소 삭제", type="primary", use_container_width=True, key="del_hotel_btn"):
                        st.session_state.confirm_delete_req = True
                        st.rerun()
            else:
                st.info("삭제할 숙소를 선택하세요.")
            
            # [요청] 튕기지 않고 바로 확인 메시지 (토글 안에서)
            if st.session_state.get('confirm_delete_req'):
                st.markdown("---")
                st.warning(f"정말 '{current_hotel}'을(를) 삭제하시겠습니까?")
                c1, c2 = st.columns(2)
                if c1.button("✅ 예"):
                    st.session_state.hotels.remove(current_hotel)
                    st.session_state.products = [p for p in st.session_state.products if p['hotel'] != current_hotel]
                    save_metadata('hotels', st.session_state.hotels)
                    save_metadata('products', st.session_state.products)
                    st.session_state.confirm_delete_req = False
                    st.session_state.pop('last_hotel', None)
                    st.success("삭제되었습니다.")
                    st.rerun()
                if c2.button("❌ 아니오"):
                    st.session_state.confirm_delete_req = False
                    st.rerun()

# ==========================================
# Main Body
# ==========================================
if current_hotel:
    st.header(f"🏨 {current_hotel} 관리")
    tab1, tab2, tab3 = st.tabs(["1. 📦 상품 세팅", "2. 📅 가격/재고 등록", "3. 📤 엑셀 추출"])

    # TAB 1: Product Setting
    with tab1:
        c1, c2 = st.columns([1, 1.5], gap="large")
        with c1:
            st.subheader("상품 등록")
            with st.form("add_p"):
                new_p = st.text_input("객실타입명 (필수)")
                new_c = st.text_input("상품관리코드 (선택사항)")
                if st.form_submit_button("추가"):
                    my_p = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
                    if new_p and new_p not in my_p:
                        st.session_state.products.append({
                            'hotel': current_hotel, 'name': new_p, 'code': new_c
                        })
                        save_metadata('products', st.session_state.products)
                        st.success("완료")
                        time.sleep(0.5)
                        st.rerun()
                    elif new_p in my_p: st.warning("이미 존재함")
        with c2:
            st.subheader("상품 순서 관리")
            curr_p_objs = [p for p in st.session_state.products if p['hotel'] == current_hotel]
            if curr_p_objs:
                for i, p in enumerate(curr_p_objs):
                    ca, cb, cc, cd = st.columns([0.5, 0.5, 4, 0.5])
                    with ca: 
                        if i>0 and st.button("⬆️", key=f"u{i}"): move_product(current_hotel, i, -1); st.rerun()
                    with cb:
                        if i<len(curr_p_objs)-1 and st.button("⬇️", key=f"d{i}"): move_product(current_hotel, i, 1); st.rerun()
                    with cc: 
                        code_txt = f" <span style='color:#888; font-size:0.8em'>({p.get('code')})</span>" if p.get('code') else ""
                        st.markdown(f"**{p['name']}**{code_txt}", unsafe_allow_html=True)
                    with cd:
                        if st.button("🗑️", key=f"del{i}"): delete_product_item(current_hotel, i); st.rerun()
                    st.divider()
            else: st.info("등록된 상품이 없습니다.")

    # TAB 2: Data Entry
    with tab2:
        with st.expander("⚡️ 데이터 일괄 생성 (클릭)", expanded=True):
            my_p_names = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
            
            if not my_p_names: st.warning("상품부터 등록해주세요.")
            else:
                c_d1, c_d2 = st.columns([1, 2])
                with c_d1:
                    dr = st.date_input("기간", [], help="시작~종료일")
                    if len(dr)==2: 
                        st.markdown(f"<div class='selected-date-box'>선택된 기간: {format_date_kr(dr[0])} ~ {format_date_kr(dr[1])}</div>", unsafe_allow_html=True)
                    else:
                        st.info("기간을 선택해주세요.")
                with c_d2:
                    st.write("요일")
                    d_lbls = ["일","월","화","수","목","금","토"]
                    sel_ds = []
                    cols = st.columns([1]*7 + [10])
                    for i, l in enumerate(d_lbls):
                        if cols[i].checkbox(l, True, key=f"wd{i}"): sel_ds.append(6 if i==0 else i-1)
                
                # [요청] 기간 추가 버튼: 검정색 (CSS 클래스 적용)
                st.markdown("""
                <style>
                /* Apply styles to the button inside the specific column/container if possible. 
                   Since we can't easily add classes to buttons, we use a CSS selector based on tree structure or assume custom component.
                   Here we use a container class wrapper. */
                .black-button button {
                    background-color: #212121 !important;
                    color: white !important;
                    border: none !important;
                }
                .black-button button:hover {
                    background-color: #424242 !important;
                }
                </style>
                """, unsafe_allow_html=True)
                
                col_btn, _ = st.columns([2, 5]) # 너비 조절
                with col_btn:
                    # div wrapper for styling
                    st.markdown('<div class="black-button">', unsafe_allow_html=True)
                    if st.button("⬇️ 기간 추가 (필수) ⬇️", key="add_pd_btn", use_container_width=True):
                        if len(dr)==2 and sel_ds:
                            nds = generate_dates(dr[0], dr[1], sel_ds)
                            if nds:
                                buf = set(st.session_state.selected_dates_buffer)
                                for d in nds: buf.add(format_date_kr(d))
                                st.session_state.selected_dates_buffer = sorted(list(buf))
                                st.rerun()
                            else: st.warning("해당 요일 없음")
                        else: st.error("기간과 요일을 모두 선택해주세요.")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                if st.session_state.selected_dates_buffer:
                    st.markdown("---")
                    upd_dates = st.multiselect("적용 날짜 확인 (삭제 가능)", st.session_state.selected_dates_buffer, st.session_state.selected_dates_buffer)
                    if len(upd_dates) != len(st.session_state.selected_dates_buffer):
                        st.session_state.selected_dates_buffer = upd_dates
                        st.rerun()
                
                st.markdown("---")
                sel_works = st.multiselect("상품 선택", my_p_names, my_p_names)
                
                input_map = {}
                reset_k = st.session_state.input_reset_key
                
                for p in sel_works:
                    st.markdown(f"🔹 {p}")
                    c1, c2, c3 = st.columns(3)
                    pr = c1.number_input("요금", key=f"pr_{p}_{reset_k}", step=1000, value=None, placeholder="숫자 입력")
                    stk = c2.number_input("재고", key=f"stk_{p}_{reset_k}", value=5)
                    stat = c3.selectbox("상태", ["Y","N"], key=f"stat_{p}_{reset_k}")
                    input_map[p] = {'p': pr, 's': stk, 'st': stat}
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # [요청] 주황색 버튼 (Primary)
                if st.button("데이터 입력하기 (저장)", type="primary", use_container_width=True):
                    if not st.session_state.selected_dates_buffer: st.error("날짜를 추가해주세요.")
                    elif not sel_works: st.error("상품을 선택해주세요.")
                    else:
                        miss = False
                        for p, v in input_map.items():
                            if v['p'] is None: miss = True
                        if miss: st.error("요금을 입력해주세요.")
                        else:
                            final_ds = []
                            for s in st.session_state.selected_dates_buffer:
                                try: final_ds.append(datetime.strptime(s.split()[0], "%Y-%m-%d").date())
                                except: pass
                            
                            new_rows = []
                            for d in final_ds:
                                for p, v in input_map.items():
                                    new_rows.append({
                                        '날짜': d, '숙소명': current_hotel, '상품명': p,
                                        '요금': v['p'], '재고': v['s'], '판매상태': v['st']
                                    })
                            
                            new_df = pd.DataFrame(new_rows)
                            st.session_state.main_df = pd.concat([st.session_state.main_df, new_df], ignore_index=True)
                            st.session_state.main_df['날짜'] = pd.to_datetime(st.session_state.main_df['날짜']).dt.date
                            st.session_state.main_df.drop_duplicates(subset=['날짜','숙소명','상품명'], keep='last', inplace=True)
                            st.session_state.main_df.sort_values(['날짜','상품명'], inplace=True)
                            
                            save_hotel_data(current_hotel, st.session_state.main_df)
                            st.session_state.selected_dates_buffer = [] 
                            st.session_state.input_reset_key += 1
                            
                            st.success("저장 완료")
                            time.sleep(1)
                            st.rerun()

        st.divider()
        st.markdown("### 📊 데이터 확인 및 수정")
        
        curr_p_order = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
        code_map = {p['name']: p.get('code', '') for p in st.session_state.products if p['hotel'] == current_hotel}
        
        show_df = st.session_state.main_df.copy()
        if not show_df.empty and curr_p_order:
            show_df['상품명'] = pd.Categorical(show_df['상품명'], curr_p_order, ordered=True)
            show_df = show_df.sort_values(['날짜', '상품명'])

        view = st.radio("보기 선택", ["📋 리스트 보기 (직접 수정 가능)", "🗓️ 요금 달력", "📅 재고 달력"], horizontal=True)
        
        if "리스트" in view:
            if show_df.empty: st.info("데이터 없음")
            else:
                list_view_df = show_df.copy()
                list_view_df['상품관리코드'] = list_view_df['상품명'].map(code_map)
                cols = ['날짜', '상품명', '상품관리코드', '요금', '재고', '판매상태']
                list_view_df = list_view_df[cols]

                edited = st.data_editor(
                    list_view_df,
                    column_config={
                        "날짜": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "상품명": st.column_config.TextColumn(disabled=True),
                        "상품관리코드": st.column_config.TextColumn(disabled=True),
                        "요금": st.column_config.NumberColumn(format="%d"),
                    },
                    use_container_width=True, hide_index=True
                )
                
                original_cols = ['날짜', '상품명', '요금', '재고', '판매상태']
                if not edited[original_cols].equals(show_df[original_cols]):
                    st.session_state.main_df.loc[edited.index, ['날짜','요금','재고','판매상태']] = edited[['날짜','요금','재고','판매상태']]
                    save_hotel_data(current_hotel, st.session_state.main_df)
                    st.toast("수정됨")
        else:
            is_stk = "재고" in view
            c1, c2, c3, c4, c5 = st.columns([6, 1, 2, 1, 6])
            
            with c2: 
                if st.button("◀️"): change_month(-1); st.rerun()
            with c4:
                if st.button("▶️"): change_month(1); st.rerun()
            
            y, m = st.session_state.cal_year, st.session_state.cal_month
            with c3:
                st.markdown(f"<h3 style='text-align: center; margin:0; padding:0;'>{y}년 {m}월</h3>", unsafe_allow_html=True)
            
            mask = (pd.to_datetime(show_df['날짜']).dt.year == y) & (pd.to_datetime(show_df['날짜']).dt.month == m)
            m_df = show_df[mask]
            
            calendar.setfirstweekday(calendar.SUNDAY)
            cal = calendar.monthcalendar(y, m)
            html = "<table class='calendar-table'><thead><tr>" + "".join([f"<th>{d}</th>" for d in ["일","월","화","수","목","금","토"]]) + "</tr></thead><tbody>"
            
            for week in cal:
                html += "<tr>"
                for d in week:
                    if d == 0: html += "<td class='other-month'></td>"
                    else:
                        d_obj = date(y, m, d)
                        recs = m_df[pd.to_datetime(m_df['날짜']).dt.date == d_obj]
                        if not recs.empty and curr_p_order:
                            recs['상품명'] = pd.Categorical(recs['상품명'], curr_p_order, ordered=True)
                            recs = recs.sort_values('상품명')
                        
                        cell = f"<span class='day-number'>{d}</span>"
                        for _, r in recs.iterrows():
                            nm = r['상품명']
                            if is_stk:
                                q = r['재고']
                                is_stop = (r['판매상태'] == 'N')
                                q_txt = f"{q}개" if q > 0 else "0 (품절)"
                                stop_txt = "<span class='stop-sales'>[판매중지]</span>" if is_stop else ""
                                cls = "stock-zero" if q == 0 else "stock-tag"
                                cell += f"<div class='prod-item'>{nm}<br>{stop_txt}<span class='{cls}'>{q_txt}</span></div>"
                            else:
                                txt, cls = f"{r['요금']:,}", "price-tag"
                                cell += f"<div class='prod-item'>{nm}<br><span class='{cls}'>{txt}</span></div>"
                        html += f"<td>{cell}</td>"
                html += "</tr>"
            html += "</tbody></table>"
            st.markdown(html, unsafe_allow_html=True)

    # TAB 3: Excel
    with tab3:
        st.subheader("엑셀 다운로드")
        if show_df.empty: 
            st.warning("데이터 없음")
        else:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.info(f"📊 현재 **{len(show_df)}**개의 데이터가 준비되었습니다. (업데이트: {now_str})")
            
            out_rows = []
            for _, r in show_df.iterrows():
                row = [""]*13
                row[0] = format_date_kr(r['날짜'])
                row[1] = r['상품명']
                row[2] = code_map.get(r['상품명'], '')
                row[6] = r['요금']; row[8] = r['재고']; row[12] = r['판매상태']
                out_rows.append(row)
            
            df_ex = pd.DataFrame(out_rows, columns=["날짜(A)", "상품명(B)", "상품코드(C)", "D", "E", "F", "요금(G)", "H", "재고(I)", "J", "K", "L", "판매상태(M)"])
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as w: df_ex.to_excel(w, index=False, sheet_name='Sheet1')
            output.seek(0)
            
            if st.download_button("📥 엑셀 파일 다운로드 (.xlsx)", output, f"[{current_hotel}]_{date.today()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", on_click=update_download_log):
                pass
            
            if st.session_state.download_logs:
                st.write("📜 다운로드 기록 (최신순)")
                for log in st.session_state.download_logs:
                    st.caption(f"✅ {log}")

else:
    st.info("👈 왼쪽에서 숙소를 선택하거나 새로 추가해주세요.")

import streamlit as st
import pandas as pd
from datetime import timedelta, date, datetime
import calendar
import io

# [추가] 드래그 앤 드롭 기능을 위한 라이브러리 임포트
try:
    from streamlit_sortables import sort_items
except ImportError:
    st.error("드래그 기능을 사용하려면 라이브러리 설치가 필요합니다. 터미널에 `pip install streamlit-sortables`를 입력해주세요.")
    def sort_items(items, key=None): return items

# --- 페이지 기본 설정 ---
st.set_page_config(layout="wide", page_title="호텔 상품 관리 시스템 Final")

# --- 스타일 커스텀 (강력한 주황색 테마 적용) ---
st.markdown("""
    <style>
    /* 1. 일반 버튼 (Secondary) */
    .stButton>button[kind="secondary"] {
        color: #e65100 !important; 
        border-color: #ffcc80 !important; 
        background-color: white !important;
    }
    .stButton>button[kind="secondary"]:hover {
        border-color: #e65100 !important;
        color: #e65100 !important;
        background-color: #fff3e0 !important;
    }

    /* 2. 주요 버튼 (Primary) */
    .stButton>button[kind="primary"] {
        background-color: #ef6c00 !important; 
        border-color: #ef6c00 !important;
        color: white !important;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #e65100 !important;
        border-color: #e65100 !important;
    }
    .stButton>button[kind="primary"]:focus {
        background-color: #e65100 !important;
        border-color: #e65100 !important;
        box-shadow: none !important;
    }

    /* 3. 탭 메뉴 선택 색상 */
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
        border-top-color: #ef6c00 !important;
        color: #ef6c00 !important;
    }

    /* 4. 캘린더 스타일 */
    .calendar-table {
        width: 100%;
        border-collapse: collapse;
    }
    .calendar-table th {
        background-color: #fff3e0;
        padding: 10px;
        text-align: center;
        border: 1px solid #ddd;
        color: #e65100;
    }
    .calendar-table td {
        vertical-align: top;
        height: 120px;
        border: 1px solid #ddd;
        padding: 5px;
        width: 14%;
    }
    .day-number {
        font-weight: bold;
        margin-bottom: 5px;
        display: block;
        color: #555;
    }
    .prod-item {
        font-size: 0.8em;
        background-color: #fff8e1;
        margin-bottom: 2px;
        padding: 2px 4px;
        border-radius: 3px;
        color: #bf360c;
        border: 1px solid #ffe0b2;
    }
    
    /* 가격 태그 (주황색) */
    .price-tag {
        font-weight: bold;
        color: #ef6c00;
    }
    
    /* 재고 태그 (기본: 파란색) */
    .stock-tag {
        font-weight: bold;
        color: #1565c0;
        background-color: #e3f2fd;
        padding: 1px 4px;
        border-radius: 4px;
        font-size: 0.9em;
    }

    /* 재고 0 (품절) 태그 - 빨간색 강조 */
    .stock-zero {
        font-weight: bold;
        color: #b71c1c; 
        background-color: #ffcdd2; 
        border: 1px solid #ef9a9a;
        padding: 1px 4px;
        border-radius: 4px;
        font-size: 0.9em;
    }

    .other-month {
        background-color: #f9f9f9;
        color: #ccc;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 데이터 저장소 (Session State) ---
if 'hotels' not in st.session_state:
    st.session_state.hotels = ["쏠비치 삼척", "소노벨 천안"]
if 'products' not in st.session_state:
    st.session_state.products = [
        {'hotel': '쏠비치 삼척', 'name': '[3인] 패밀리 스탠다드'},
        {'hotel': '쏠비치 삼척', 'name': '[4인] 스위트 오션'}
    ]
if 'main_df' not in st.session_state:
    st.session_state.main_df = pd.DataFrame(columns=['날짜', '숙소명', '상품명', '요금', '재고', '판매상태'])

# 캘린더 보기용 현재 월 상태
if 'cal_year' not in st.session_state:
    st.session_state.cal_year = date.today().year
if 'cal_month' not in st.session_state:
    st.session_state.cal_month = date.today().month

# 삭제 확인용
if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False
if 'hotel_to_delete' not in st.session_state:
    st.session_state.hotel_to_delete = None

# --- 편의 함수 ---
def format_date_kr(d):
    """YYYY-MM-DD (요일) 형식으로 변환"""
    # 안전하게 datetime.date 객체로 변환
    if isinstance(d, str):
        d = pd.to_datetime(d).date()
    elif isinstance(d, datetime):
        d = d.date()
        
    weekdays = ["월", "화", "수", "목", "금", "토", "일"]
    # [수정] 포맷 변경: YYYY/MM/DD -> YYYY-MM-DD
    return f"{d.year}-{d.month:02d}-{d.day:02d} ({weekdays[d.weekday()]})"

def generate_dates(start_date, end_date, weekdays):
    dates = []
    current_date = start_date
    while current_date <= end_date:
        # python weekday: 0=Mon, 6=Sun
        if current_date.weekday() in weekdays:
            dates.append(current_date)
        current_date += timedelta(days=1)
    return dates

def change_month(amount):
    st.session_state.cal_month += amount
    if st.session_state.cal_month > 12:
        st.session_state.cal_month = 1
        st.session_state.cal_year += 1
    elif st.session_state.cal_month < 1:
        st.session_state.cal_month = 12
        st.session_state.cal_year -= 1

# ==========================================
# 1. 사이드바: 숙소 선택 및 관리
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
                new_hotel_input = st.text_input("새 숙소명 입력")
                if st.form_submit_button("➕ 추가하기"):
                    if new_hotel_input and new_hotel_input not in st.session_state.hotels:
                        st.session_state.hotels.append(new_hotel_input)
                        st.success(f"'{new_hotel_input}' 추가되었습니다.")
                        st.rerun()
                    elif new_hotel_input in st.session_state.hotels:
                        st.warning("이미 존재하는 숙소입니다.")
        with tab_del:
            del_search = st.text_input("삭제할 숙소 찾기")
            del_candidates = [h for h in st.session_state.hotels if del_search in h]
            if del_candidates:
                target_del = st.radio("삭제 대상 선택", del_candidates)
                if st.button("🗑 선택한 숙소 삭제"):
                    st.session_state.confirm_delete = True
                    st.session_state.hotel_to_delete = target_del
                if st.session_state.confirm_delete:
                    st.error(f"정말 '{st.session_state.hotel_to_delete}'를 삭제하시겠습니까?")
                    c_y, c_n = st.columns(2)
                    if c_y.button("✅ 예"):
                        target = st.session_state.hotel_to_delete
                        st.session_state.hotels.remove(target)
                        st.session_state.products = [p for p in st.session_state.products if p['hotel'] != target]
                        st.session_state.main_df = st.session_state.main_df[st.session_state.main_df['숙소명'] != target]
                        st.session_state.confirm_delete = False
                        st.session_state.hotel_to_delete = None
                        st.success("삭제되었습니다.")
                        st.rerun()
                    if c_n.button("❌ 아니오"):
                        st.session_state.confirm_delete = False
                        st.rerun()

# ==========================================
# 2. 메인 작업 영역
# ==========================================
if current_hotel:
    st.header(f"🏨 {current_hotel} 관리")
    
    tab_prod, tab_work, tab_excel = st.tabs(["1. 📦 상품 세팅", "2. 📅 가격/재고 등록 & 확인", "3. 📤 엑셀 추출"])

    # ---------------------------------------------------
    # TAB 1: 상품 세팅
    # ---------------------------------------------------
    with tab_prod:
        c1, c2 = st.columns([1, 2])
        with c1:
            st.subheader("상품 등록")
            with st.form("add_product_form", clear_on_submit=True):
                new_prod_name = st.text_input("상품명 (객실타입) 입력")
                if st.form_submit_button("상품 추가"):
                    my_prods = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
                    if new_prod_name and new_prod_name not in my_prods:
                        st.session_state.products.append({'hotel': current_hotel, 'name': new_prod_name})
                        st.success(f"'{new_prod_name}' 추가 완료")
                        st.rerun()
                    elif new_prod_name in my_prods:
                        st.warning("이미 등록된 상품명입니다.")
        with c2:
            st.subheader("등록된 상품 리스트 (순서 변경 가능)")
            st.caption("💡 리스트를 드래그하여 위아래로 순서를 바꿀 수 있습니다.")
            current_products_objs = [p for p in st.session_state.products if p['hotel'] == current_hotel]
            current_names = [p['name'] for p in current_products_objs]
            if current_names:
                sorted_names = sort_items(current_names)
                if sorted_names != current_names:
                    other_hotel_products = [p for p in st.session_state.products if p['hotel'] != current_hotel]
                    new_ordered_products = [{'hotel': current_hotel, 'name': name} for name in sorted_names]
                    st.session_state.products = other_hotel_products + new_ordered_products
                    st.rerun() 
                st.markdown("###### 📋 확정된 순서")
                df_show = pd.DataFrame(sorted_names, columns=["상품명"])
                df_show.index = df_show.index + 1  
                st.table(df_show)
            else:
                st.info("등록된 상품이 없습니다. 왼쪽에서 상품을 추가해주세요.")

    # ---------------------------------------------------
    # TAB 2: 가격/재고 등록 및 뷰어
    # ---------------------------------------------------
    with tab_work:
        # [일괄 입력 패널]
        with st.expander("⚡️ 가격/재고 일괄 입력 열기 (클릭)", expanded=True):
            my_products = [p['name'] for p in st.session_state.products if p['hotel'] == current_hotel]
            if not my_products:
                st.warning("상품을 먼저 등록해주세요.")
            else:
                st.markdown("#### 1. 기간 및 요일")
                c_d1, c_d2 = st.columns([1, 2])
                with c_d1: d_range = st.date_input("기간", [])
                with c_d2:
                    st.markdown("요일 선택")
                    ui_labels = ["일", "월", "화", "수", "목", "금", "토"]
                    py_weekdays = [6, 0, 1, 2, 3, 4, 5]
                    sel_days = []
                    day_cols = st.columns(7) 
                    for i, label in enumerate(ui_labels):
                        if day_cols[i].checkbox(label, value=True, key=f"day_chk_{i}"):
                            sel_days.append(py_weekdays[i])
                
                st.markdown("#### 2. 상품별 설정")
                sel_work_prods = st.multiselect("작업할 상품 선택", my_products, default=my_products)
                
                input_map = {}
                for p in sel_work_prods:
                    st.markdown(f"**🔹 {p}**") 
                    pc1, pc2, pc3 = st.columns(3)
                    pr = pc1.number_input(f"요금 (원)", key=f"p_{p}", value=None, step=1000, placeholder="숫자 입력")
                    stk = pc2.number_input(f"재고 (개)", key=f"s_{p}", value=5)
                    sts = pc3.selectbox(f"상태", ["Y", "N"], key=f"st_{p}")
                    input_map[p] = {'price': pr, 'stock': stk, 'status': sts}
                
                if st.button("💾 데이터 생성하기", type="primary", use_container_width=True):
                    if len(d_range) != 2:
                        st.error("🚨 날짜(기간)를 정확히 선택해주세요.")
                    elif not sel_days:
                        st.error("🚨 요일을 최소 하나 이상 선택해주세요.")
                    elif not sel_work_prods:
                        st.error("🚨 작업할 상품을 선택해주세요.")
                    else:
                        missing_price = False
                        for p, val in input_map.items():
                            if val['price'] is None:
                                missing_price = True; break
                        if missing_price:
                            st.error("🚨 모든 상품의 요금을 입력해주세요.")
                        else:
                            dates = generate_dates(d_range[0], d_range[1], sel_days)
                            if not dates:
                                st.warning("선택된 기간 내에 해당 요일이 존재하지 않습니다.")
                            else:
                                new_rows = []
                                for d in dates:
                                    for p, val in input_map.items():
                                        new_rows.append({
                                            '날짜': d, '숙소명': current_hotel, '상품명': p,
                                            '요금': val['price'], '재고': val['stock'], '판매상태': val['status']
                                        })
                                st.session_state.main_df = pd.concat([st.session_state.main_df, pd.DataFrame(new_rows)], ignore_index=True)
                                st.session_state.main_df['날짜'] = pd.to_datetime(st.session_state.main_df['날짜']).dt.date
                                st.session_state.main_df.drop_duplicates(subset=['날짜', '숙소명', '상품명'], keep='last', inplace=True)
                                st.session_state.main_df.sort_values(['날짜', '상품명'], inplace=True)
                                st.success(f"{len(new_rows)}건 생성 완료!")
                                st.rerun()

        # [데이터 확인 및 수정 영역]
        st.divider()
        st.markdown("### 📊 데이터 확인 및 수정")
        
        hotel_df = st.session_state.main_df[st.session_state.main_df['숙소명'] == current_hotel].copy()
        
        view_type = st.radio(
            "보기 방식", 
            ["📋 리스트 표보기 (직접 수정 가능)", "🗓️ 월별 요금 캘린더 보기", "📅 월별 재고 캘린더 보기"], 
            horizontal=True
        )
        
        # 1) 리스트 표 보기 (직접 수정)
        if view_type == "📋 리스트 표보기 (직접 수정 가능)":
            if not hotel_df.empty:
                # [수정] 날짜 포맷 변경 (YYYY-MM-DD)
                # 에디터 특성상 YYYY-MM-DD 형식을 써야 달력 기능을 유지할 수 있습니다.
                column_config = {
                    "날짜": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
                    "상품명": st.column_config.TextColumn("상품명", disabled=True),
                    "요금": st.column_config.NumberColumn("요금", format="%d", step=1000), 
                    "재고": st.column_config.NumberColumn("재고", step=1),
                    "판매상태": st.column_config.SelectboxColumn("판매상태", options=["Y", "N"])
                }
                
                if not hotel_df.empty:
                    hotel_df['요금'] = pd.to_numeric(hotel_df['요금'], errors='coerce').fillna(0).astype(int)

                # 에디터 표시
                edited_df = st.data_editor(
                    hotel_df[['날짜', '상품명', '요금', '재고', '판매상태']], 
                    column_config=column_config,
                    use_container_width=True, 
                    hide_index=True,
                    key="list_editor"
                )
                # 수정 저장
                if not edited_df.equals(hotel_df[['날짜', '상품명', '요금', '재고', '판매상태']]):
                    st.session_state.main_df.loc[edited_df.index, ['날짜', '요금', '재고', '판매상태']] = edited_df[['날짜', '요금', '재고', '판매상태']]
                    st.toast("✅ 수정사항이 저장되었습니다!")
            else:
                st.info("데이터가 없습니다. 위에서 데이터를 생성해주세요.")
                
        # 2) 캘린더 보기 (요금 or 재고)
        else:
            is_stock_view = "재고" in view_type
            
            # [상단] 월 이동 & 타이틀
            _, c_prev, c_title, c_next, _ = st.columns([5, 0.5, 2, 0.5, 5])
            
            with c_prev:
                if st.button("⬅️"): change_month(-1); st.rerun()
            with c_next:
                if st.button("➡️>"): change_month(1); st.rerun()
            with c_title:
                curr_y = st.session_state.cal_year
                curr_m = st.session_state.cal_month
                st.markdown(f"<h3 style='text-align: center; color: #e65100; margin-top: -5px; white-space: nowrap;'>{curr_y}년 {curr_m}월</h3>", unsafe_allow_html=True)
            
            # 데이터 필터링
            hotel_df['날짜'] = pd.to_datetime(hotel_df['날짜'])
            mask = (hotel_df['날짜'].dt.year == curr_y) & (hotel_df['날짜'].dt.month == curr_m)
            month_data = hotel_df[mask].copy()
            
            # [A] 비주얼 캘린더 (HTML)
            cal = calendar.monthcalendar(curr_y, curr_m)
            html_cal = "<table class='calendar-table'><thead><tr>" + "".join([f"<th>{d}</th>" for d in ["월", "화", "수", "목", "금", "토", "일"]]) + "</tr></thead><tbody>"
            for week in cal:
                html_cal += "<tr>"
                for day in week:
                    if day == 0: html_cal += "<td class='other-month'></td>"
                    else:
                        d_obj = date(curr_y, curr_m, day)
                        day_records = month_data[month_data['날짜'].dt.date == d_obj]
                        cell_content = f"<span class='day-number'>{day}</span>"
                        for _, row in day_records.iterrows():
                            p_name = row['상품명']
                            if is_stock_view:
                                qty = row['재고']
                                if qty == 0:
                                    p_val = "0개 (품절)"; css_class = "stock-zero"
                                else:
                                    p_val = f"{qty}개"; css_class = "stock-tag"
                            else:
                                p_val = f"{row['요금']:,}원"; css_class = "price-tag"
                            cell_content += f"<div class='prod-item'>{p_name}<br><span class='{css_class}'>{p_val}</span></div>"
                        html_cal += f"<td>{cell_content}</td>"
                html_cal += "</tr>"
            html_cal += "</tbody></table>"
            st.markdown(html_cal, unsafe_allow_html=True)

    # ---------------------------------------------------
    # TAB 3: 엑셀 추출
    # ---------------------------------------------------
    with tab_excel:
        st.subheader("업로드용 엑셀 다운로드")
        
        hotel_df = st.session_state.main_df[st.session_state.main_df['숙소명'] == current_hotel].copy()
        
        if hotel_df.empty:
            st.warning("⚠️ 엑셀로 추출할 데이터가 없습니다. 먼저 '상품 세팅' 및 '가격/재고 등록'을 진행해주세요.")
        else:
            st.success(f"총 {len(hotel_df)}개의 데이터가 준비되었습니다.")
            
            # 엑셀 파일 생성 로직 (메모리 버퍼 사용)
            export_data = []
            for _, row in hotel_df.iterrows():
                r = [""] * 13
                # [수정] 날짜 포맷팅: YYYY-MM-DD (요일)
                try:
                    r[0] = format_date_kr(row['날짜'])
                except Exception:
                    r[0] = str(row['날짜'])
                    
                r[1] = row['상품명']
                r[6] = row['요금']
                r[8] = row['재고']
                r[12] = row['판매상태']
                export_data.append(r)
            
            df_ex = pd.DataFrame(export_data, columns=["날짜(A)", "상품명(B)", "C", "D", "E", "F", "요금(G)", "H", "재고(I)", "J", "K", "L", "판매상태(M)"])
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)
            
            st.download_button(
                label="📥 엑셀 파일 다운로드 (.xlsx)",
                data=output,
                file_name=f"[{current_hotel}]_upload_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

else:
    st.info("왼쪽 사이드바에서 숙소를 검색하거나 선택해주세요.")
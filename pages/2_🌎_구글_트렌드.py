import streamlit as st
import pandas as pd
from pytrends.request import TrendReq
import matplotlib.pyplot as plt
from datetime import datetime, date
import io
import time
import random

# 1. 페이지 설정
st.set_page_config(page_title="글로벌 트렌드 분석 도구", layout="wide")
st.title("🌎 Global Google Trends Dashboard")
st.markdown("### 오류가 수정된 안전 모드 트렌드 분석기")

# 2. 구글 트렌드 설정
def get_pytrends():
    return TrendReq(hl='ko', tz=324, retries=5, backoff_factor=2)

@st.cache_data(ttl=3600)
def fetch_google_data_safe(terms, timeframe, geo):
    time.sleep(random.uniform(8, 15)) 
    py_instance = get_pytrends()
    try:
        py_instance.build_payload(terms, timeframe=timeframe, geo=geo)
        data = py_instance.interest_over_time()
        return data
    except Exception as e:
        if "429" in str(e):
            return "BLOCK"
        raise e

# 3. 사이드바 구성
with st.sidebar:
    st.header("📁 데이터 관리")
    uploaded_file = st.file_uploader("분석할 엑셀 파일을 업로드하세요", type=["xlsx"])
    
    st.divider()
    st.subheader("⚙️ 분석 조건")
    country_map = {"대한민국": "KR", "대만": "TW", "싱가포르": "SG", "몽골": "MN", "전 세계": ""}
    selected_country = st.selectbox("분석 국가", list(country_map.keys()))
    geo = country_map[selected_country]
    
    date_mode = st.radio("기간 설정 방식", ["기본 설정 사용", "직접 날짜 입력"])
    if date_mode == "기본 설정 사용":
        timeframe = st.selectbox("기간 선택", ["today 12-m", "today 3-m", "today 1-m", "today 5-y"])
    else:
        col1, col2 = st.columns(2)
        with col1: start_date = st.date_input("시작일", date(2025, 1, 1))
        with col2: end_date = st.date_input("종료일", date.today())
        timeframe = f"{start_date.strftime('%Y-%m-%d')} {end_date.strftime('%Y-%m-%d')}"

# 4. 분석 메인 로직
if uploaded_file:
    df_input = pd.read_excel(uploaded_file)
    if 'GroupName' in df_input.columns:
        all_groups = []
        for _, row in df_input.iterrows():
            g_name = str(row['GroupName']).strip()
            if not g_name or g_name.startswith('*') or g_name == "nan": continue
            kw_val = str(row['Keywords']).strip() if 'Keywords' in df_input.columns and pd.notnull(row['Keywords']) else ""
            search_term = " ".join([g_name] + ([k.strip() for k in kw_val.split(',')] if kw_val and kw_val != "nan" else []))
            all_groups.append({"name": g_name, "term": search_term})

        if all_groups:
            anchor = all_groups[0]
            others = all_groups[1:]
            
            if st.sidebar.button("🚀 분석 시작"):
                final_df = pd.DataFrame()
                reference_data = pd.DataFrame()
                status = st.empty()
                progress = st.progress(0)

                batch_size = 2 
                for i in range(0, len(others) if others else 1, batch_size):
                    chunk = others[i:i+batch_size]
                    current_names = [anchor['name']] + [c['name'] for c in chunk]
                    current_terms = [anchor['term']] + [c['term'] for c in chunk]
                    
                    status.text(f"⏳ 데이터 수집 중... ({i//batch_size + 1}번째 배치)")
                    batch_res = fetch_google_data_safe(current_terms, timeframe, geo)
                    
                    # [핵심 수정 부분] isinstance를 사용하여 타입 확인 후 비교
                    if isinstance(batch_res, str) and batch_res == "BLOCK":
                        st.error("🚨 구글 접속이 차단되었습니다. 약 30분~1시간 뒤에 다시 시도해 주세요.")
                        st.stop()
                    
                    if batch_res is not None and not batch_res.empty:
                        if 'isPartial' in batch_res.columns: batch_res = batch_res.drop(columns=['isPartial'])
                        batch_res.columns = current_names
                        
                        if i == 0:
                            reference_data = batch_res[[anchor['name']]].copy()
                            final_df = batch_res
                        else:
                            curr_anchor = batch_res[[anchor['name']]]
                            scale_factor = (reference_data[anchor['name']] / curr_anchor[anchor['name']]).fillna(1).replace([float('inf'), -float('inf')], 1)
                            for col in batch_res.columns: batch_res[col] = batch_res[col] * scale_factor
                            final_df = pd.concat([final_df, batch_res.drop(columns=[anchor['name']])], axis=1)
                    
                    progress.progress(min((i + batch_size) / (len(others) + 1) if others else 1.0, 1.0))

                status.empty()
                if not final_df.empty:
                    st.session_state['result'] = final_df
                    st.success("✅ 분석 완료!")

        if 'result' in st.session_state:
            res_df = st.session_state['result']
            st.divider()
            selected = st.multiselect("📈 표시할 키워드 선택:", options=res_df.columns.tolist(), default=res_df.columns.tolist())
            if selected:
                st.line_chart(res_df[selected])
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res_df[selected].to_excel(writer, index=True)
                st.download_button("📥 엑셀 저장", output.getvalue(), file_name="google_trend.xlsx")
                st.dataframe(res_df[selected], use_container_width=True)

import streamlit as st
import pandas as pd
from pytrends.request import TrendReq
import matplotlib.pyplot as plt
from datetime import datetime, date
import io
import time
import random
import os

# 1. 페이지 설정
st.set_page_config(page_title="글로벌 트렌드 분석 도구 | 프리덤", layout="wide")
st.title("🌎 Global Google Trends Dashboard")
st.markdown("### 오류가 수정된 안전 모드 트렌드 분석기")

# 2. 구글 트렌드 설정 함수
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
    
    # --- [추가 포인트] 분석 양식 다운로드 버튼 ---
    # 깃허브 메인 폴더 혹은 현재 폴더에서 파일을 찾습니다.
    template_file = "keywords_input_.xlsx"
    # 스트림릿 클라우드 경로 대응 (root와 pages 폴더 고려)
    paths = [template_file, os.path.join("..", template_file)]
    target_path = next((p for p in paths if os.path.exists(p)), None)

    if target_path:
        with open(target_path, "rb") as f:
            st.download_button(
                label="📥 분석 양식(Excel) 받기",
                data=f,
                file_name="trend_analysis_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="분석에 필요한 키워드 입력 양식을 다운로드합니다."
            )
    else:
        st.error("⚠️ 'keywords_input_.xlsx' 파일을 찾을 수 없습니다. 깃허브 메인 폴더를 확인해주세요.")
    
    st.divider()

    # 파일 업로드 (양식을 받은 후 수정해서 다시 넣는 곳)
    uploaded_file = st.file_uploader("수정한 엑셀 파일을 업로드하세요", type=["xlsx"])
    
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
    try:
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
                        
                        # [에러 처리] 구글 차단 시 앱 멈춤 방지
                        if isinstance(batch_res, str) and batch_res == "BLOCK":
                            st.error("🚨 구글 접속이 일시적으로 차단되었습니다. 약 30분~1시간 뒤에 다시 시도해 주세요.")
                            st.stop()
                        
                        if batch_res is not None and not batch_res.empty:
                            if 'isPartial' in batch_res.columns: batch_res = batch_res.drop(columns=['isPartial'])
                            batch_res.columns = current_names
                            
                            if i == 0:
                                reference_data = batch_res[[anchor['name']]].copy()
                                final_df = batch_res
                            else:
                                curr_anchor = batch_res[[anchor['name']]]
                                # 0으로 나누기 방지 및 스케일 조정
                                scale_factor = (reference_data[anchor['name']] / curr_anchor[anchor['name']]).fillna(1).replace([float('inf'), -float('inf')], 1)
                                for col in batch_res.columns: batch_res[col] = batch_res[col] * scale_factor
                                final_df = pd.concat([final_df, batch_res.drop(columns=[anchor['name']])], axis=1)
                        
                        progress.progress(min((i + batch_size) / (len(others) + 1) if others else 1.0, 1.0))

                    status.empty()
                    if not final_df.empty:
                        st.session_state['result'] = final_df
                        st.success("✅ 분석 완료!")
        else:
            st.error("❌ 엑셀 파일에 'GroupName' 컬럼이 없습니다. 양식을 다시 확인해 주세요.")
    except Exception as e:
        st.error(f"파일을 읽는 중 에러가 발생했습니다: {e}")

# 5. 결과 시각화 및 저장
if 'result' in st.session_state:
    res_df = st.session_state['result']
    st.divider()
    selected = st.multiselect("📈 표시할 키워드 선택:", options=res_df.columns.tolist(), default=res_df.columns.tolist())
    if selected:
        st.line_chart(res_df[selected])
        
        # 엑셀 다운로드 파일 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df[selected].to_excel(writer, index=True)
        
        st.download_button("📥 분석 결과(Excel) 저장", output.getvalue(), file_name="google_trend_result.xlsx")
        st.dataframe(res_df[selected], use_container_width=True)

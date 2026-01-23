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

# 2. 구글 트렌드 설정 함수
def get_pytrends():
    # requests_args를 추가하여 타임아웃 설정을 강화
    return TrendReq(hl='ko', tz=324, retries=3, backoff_factor=1, timeout=(10, 25))

@st.cache_data(ttl=3600)
def fetch_google_data_safe(terms, timeframe, geo):
    time.sleep(random.uniform(2, 5))  # 랜덤 대기 시간
    py_instance = get_pytrends()
    try:
        py_instance.build_payload(terms, timeframe=timeframe, geo=geo)
        data = py_instance.interest_over_time()
        return data
    except Exception as e:
        err_msg = str(e)
        if "429" in err_msg:
            return "BLOCK"
        # urllib3 버전 이슈 등으로 인한 TypeError 처리
        if "method_whitelist" in err_msg or "allowed_methods" in err_msg:
            st.error("⚠️ 서버 버전 충돌이 감지되었습니다. requirements.txt에 'urllib3<2.0.0'을 추가해주세요.")
            return None
        raise e

# 3. 사이드바 구성
with st.sidebar:
    st.header("📁 데이터 관리")
    
    # [분석 양식 다운로드 버튼]
    template_file = "keywords_input_.xlsx"
    paths = [template_file, os.path.join("..", template_file)]
    target_path = next((p for p in paths if os.path.exists(p)), None)

    if target_path:
        with open(target_path, "rb") as f:
            st.download_button(
                label="📥 분석 양식(Excel) 받기",
                data=f,
                file_name="trend_analysis_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    st.divider()

    # --- 입력 방식 선택 (직접 입력 vs 엑셀) ---
    input_mode = st.radio("입력 방식 선택", ["직접 입력 (최대 2개)", "엑셀 파일 업로드"])
    
    all_groups = [] # 분석할 데이터를 담을 리스트

    if input_mode == "직접 입력 (최대 2개)":
        st.subheader("⌨️ 키워드 직접 입력")
        st.info("💡 추가 키워드를 비워두면 '그룹 이름'만으로 검색합니다.")
        
        # 그룹 1
        g1_name = st.text_input("그룹 1 이름 (기준)", "프리덤", key="g1_n")
        g1_kws = st.text_input("그룹 1 추가 키워드 (선택)", "", key="g1_k", placeholder="비워두면 그룹명만 검색")
        
        # 그룹 2
        g2_name = st.text_input("그룹 2 이름 (비교)", "", key="g2_n", placeholder="비교할 경쟁사 이름")
        g2_kws = st.text_input("그룹 2 추가 키워드 (선택)", "", key="g2_k", placeholder="비워두면 그룹명만 검색")

        # [수정] 키워드가 비어있을 경우 처리 로직
        if g1_name:
            extra_kws = [k.strip() for k in g1_kws.split(",") if k.strip()]
            if extra_kws:
                term1 = g1_name + " " + " ".join(extra_kws)
            else:
                term1 = g1_name # 키워드 없으면 그룹명만 사용
            all_groups.append({"name": g1_name, "term": term1.strip()})
        
        if g2_name:
            extra_kws2 = [k.strip() for k in g2_kws.split(",") if k.strip()]
            if extra_kws2:
                term2 = g2_name + " " + " ".join(extra_kws2)
            else:
                term2 = g2_name # 키워드 없으면 그룹명만 사용
            all_groups.append({"name": g2_name, "term": term2.strip()})

    else:
        # 엑셀 파일 업로드 모드
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            try:
                df_input = pd.read_excel(uploaded_file)
                
                # 컬럼명 찾기 (대소문자 무관)
                cols = {c.lower(): c for c in df_input.columns}
                name_col_key = next((k for k in cols if k in ['groupname', '그룹명']), None)
                kw_col_key = next((k for k in cols if k in ['keywords', '키워드']), None)

                if name_col_key:
                    real_name_col = cols[name_col_key]
                    
                    for _, row in df_input.iterrows():
                        g_name = str(row[real_name_col]).strip()
                        if not g_name or g_name.startswith('*') or g_name == "nan": continue
                        
                        # [수정] 엑셀 키워드 컬럼 처리 강화
                        search_term = g_name # 기본값은 그룹명
                        
                        if kw_col_key:
                            real_kw_col = cols[kw_col_key]
                            raw_val = row[real_kw_col]
                            
                            # 값이 존재하고 nan이 아닐 때만 키워드 병합
                            if pd.notnull(raw_val) and str(raw_val).lower() != 'nan' and str(raw_val).strip() != '':
                                kw_list = [k.strip() for k in str(raw_val).split(',') if k.strip()]
                                if kw_list:
                                    search_term = g_name + " " + " ".join(kw_list)
                        
                        all_groups.append({"name": g_name, "term": search_term})
                else:
                    st.error("엑셀에 'GroupName(또는 그룹명)' 컬럼이 반드시 필요합니다.")
            except Exception as e:
                st.error(f"파일 처리 중 오류가 발생했습니다: {e}")

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
if all_groups:
    anchor = all_groups[0]
    others = all_groups[1:]
    
    st.markdown(f"**📊 분석 대상:** {' vs '.join([g['name'] for g in all_groups])}")

    if st.button("🚀 분석 시작 (Run Analysis)"):
        final_df = pd.DataFrame()
        reference_data = pd.DataFrame()
        status = st.empty()
        progress = st.progress(0)

        # 배치 사이즈 설정
        batch_size = 4 
        loop_target = others if others else []
        
        # 1. 단일 그룹 분석 (비교 대상 없을 때)
        if not loop_target:
             status.text(f"⏳ 데이터 수집 중... (단일 그룹)")
             batch_res = fetch_google_data_safe([anchor['term']], timeframe, geo)
             
             if isinstance(batch_res, str) and batch_res == "BLOCK":
                st.error("🚨 구글 접속이 일시 차단되었습니다. 잠시 후 다시 시도해주세요.")
             elif batch_res is not None and not batch_res.empty:
                 if 'isPartial' in batch_res.columns: batch_res = batch_res.drop(columns=['isPartial'])
                 batch_res.columns = [anchor['name']]
                 final_df = batch_res
                 st.session_state['result'] = final_df
                 st.success("✅ 분석 완료!")
             else:
                 st.warning(f"데이터가 없습니다. 검색어('{anchor['term']}')를 확인해주세요.")

        # 2. 다중 그룹 분석 (비교 대상 있을 때)
        else:
            for i in range(0, len(loop_target), batch_size):
                chunk = loop_target[i:i+batch_size]
                current_names = [anchor['name']] + [c['name'] for c in chunk]
                current_terms = [anchor['term']] + [c['term'] for c in chunk]
                
                status.text(f"⏳ 데이터 수집 중... ({i//batch_size + 1}번째 배치)")
                batch_res = fetch_google_data_safe(current_terms, timeframe, geo)
                
                if isinstance(batch_res, str) and batch_res == "BLOCK":
                    st.error("🚨 구글 접속이 일시적으로 차단되었습니다. 30분 뒤 재시도해주세요.")
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
                
                progress.progress(min((i + batch_size) / (len(loop_target) + 1), 1.0))

            status.empty()
            if not final_df.empty:
                st.session_state['result'] = final_df
                st.success("✅ 분석 완료!")

# 5. 결과 시각화 및 저장
if 'result' in st.session_state:
    res_df = st.session_state['result']
    st.divider()
    
    selected = st.multiselect("📈 표시할 키워드 선택:", options=res_df.columns.tolist(), default=res_df.columns.tolist())
    
    if selected:
        st.line_chart(res_df[selected])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df[selected].to_excel(writer, index=True)
        
        st.download_button("📥 분석 결과(Excel) 저장", output.getvalue(), file_name="google_trend_result.xlsx")
        st.dataframe(res_df[selected], use_container_width=True)
elif not all_groups:
    st.info("👈 사이드바에서 키워드를 입력하고 '분석 시작'을 눌러주세요.")

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
    time.sleep(random.uniform(5, 10))  # 요청 간격을 조금 넉넉하게 둠
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
    else:
        st.warning("⚠️ 양식 파일(keywords_input_.xlsx)을 찾을 수 없습니다.")
    
    st.divider()

    # --- [수정 포인트] 입력 방식 선택 (직접 입력 vs 엑셀) ---
    input_mode = st.radio("입력 방식 선택", ["직접 입력 (최대 2개)", "엑셀 파일 업로드"])
    
    all_groups = [] # 분석할 데이터를 담을 리스트

    if input_mode == "직접 입력 (최대 2개)":
        st.subheader("⌨️ 키워드 직접 입력")
        st.caption("그룹명과 추가 키워드는 띄어쓰기로 합쳐져서 검색됩니다. (예: 프리덤 짐웨어)")
        
        # 그룹 1
        g1_name = st.text_input("그룹 1 이름 (기준)", "프리덤", key="g1_n")
        g1_kws = st.text_input("그룹 1 추가 키워드", "짐웨어, 애슬레저", key="g1_k")
        
        # 그룹 2
        g2_name = st.text_input("그룹 2 이름 (비교)", "나이키", key="g2_n")
        g2_kws = st.text_input("그룹 2 추가 키워드", "운동화, 런닝", key="g2_k")

        # 데이터 변환 로직
        if g1_name:
            term1 = g1_name + " " + " ".join([k.strip() for k in g1_kws.split(",") if k.strip()])
            all_groups.append({"name": g1_name, "term": term1.strip()})
        
        if g2_name:
            term2 = g2_name + " " + " ".join([k.strip() for k in g2_kws.split(",") if k.strip()])
            all_groups.append({"name": g2_name, "term": term2.strip()})

    else:
        # 엑셀 파일 업로드 모드
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            try:
                df_input = pd.read_excel(uploaded_file)
                if 'GroupName' in df_input.columns:
                    for _, row in df_input.iterrows():
                        g_name = str(row['GroupName']).strip()
                        if not g_name or g_name.startswith('*') or g_name == "nan": continue
                        kw_val = str(row['Keywords']).strip() if 'Keywords' in df_input.columns and pd.notnull(row['Keywords']) else ""
                        search_term = " ".join([g_name] + ([k.strip() for k in kw_val.split(',')] if kw_val and kw_val != "nan" else []))
                        all_groups.append({"name": g_name, "term": search_term})
                else:
                    st.error("엑셀에 'GroupName' 컬럼이 필요합니다.")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

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
    
    # 데이터 미리보기
    st.write(f"📊 **분석 대상:** {', '.join([g['name'] for g in all_groups])}")

    if st.button("🚀 분석 시작 (Run Analysis)"):
        final_df = pd.DataFrame()
        reference_data = pd.DataFrame()
        status = st.empty()
        progress = st.progress(0)

        # 배치 사이즈 (구글은 한번에 5개까지 가능하므로 2개 입력 시에는 한 번에 처리됨)
        batch_size = 4 
        
        # others가 없어도(그룹이 1개여도) 돌아가도록 로직 처리
        loop_target = others if others else []
        
        # 만약 비교 그룹이 없으면(그룹 1개만 입력 시) 단독 실행
        if not loop_target:
             status.text(f"⏳ 데이터 수집 중... (단일 그룹)")
             batch_res = fetch_google_data_safe([anchor['term']], timeframe, geo)
             if isinstance(batch_res, str) and batch_res == "BLOCK":
                st.error("🚨 구글 접속이 일시 차단되었습니다. 잠시 후 다시 시도해주세요.")
                st.stop()
             if batch_res is not None and not batch_res.empty:
                 if 'isPartial' in batch_res.columns: batch_res = batch_res.drop(columns=['isPartial'])
                 batch_res.columns = [anchor['name']]
                 final_df = batch_res
                 st.session_state['result'] = final_df
                 st.success("✅ 분석 완료!")
        else:
            # 여러 그룹일 경우 배치 처리
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
                        # 스케일링 로직 (기준 그룹을 통해 데이터 보정)
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
    
    # 키워드 선택 (기본적으로 전체 선택)
    selected = st.multiselect("📈 표시할 키워드 선택:", options=res_df.columns.tolist(), default=res_df.columns.tolist())
    
    if selected:
        st.line_chart(res_df[selected])
        
        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df[selected].to_excel(writer, index=True)
        
        st.download_button("📥 분석 결과(Excel) 저장", output.getvalue(), file_name="google_trend_result.xlsx")
        st.dataframe(res_df[selected], use_container_width=True)
elif not all_groups:
    st.info("👈 사이드바에서 '직접 입력'을 하거나 '엑셀 파일'을 업로드해주세요.")

import streamlit as st
import pandas as pd
import requests
import json
from datetime import datetime
import io
import os

# 1. 페이지 설정
st.set_page_config(page_title="프리덤 네이버 트렌드 | 자동 키워드", layout="wide")
st.title("🏃‍♂️ Freedom Naver Trend Dashboard")

# 2. 보안 설정 (Secrets)
try:
    CLIENT_ID = st.secrets["NAVER_CLIENT_ID"]
    CLIENT_SECRET = st.secrets["NAVER_CLIENT_SECRET"]
except KeyError:
    st.error("오류: Streamlit Secrets에서 NAVER_CLIENT_ID와 SECRET을 설정해주세요.")
    st.stop()

# 연령대 매핑 딕셔너리
AGE_MAP = {
    "0~12세": "1", "13~18세": "2", "19~24세": "3", "25~29세": "4",
    "30~34세": "5", "35~39세": "6", "40~44세": "7", "45~49세": "8",
    "50~54세": "9", "55~59세": "10", "60세 이상": "11"
}

# 3. Naver API 호출 함수
def get_api_data(keyword_groups, gender, age_codes):
    url = "https://openapi.naver.com/v1/datalab/search"
    headers = {
        "X-Naver-Client-Id": CLIENT_ID,
        "X-Naver-Client-Secret": CLIENT_SECRET,
        "Content-Type": "application/json"
    }
    body = {
        "startDate": "2024-01-01",
        "endDate": datetime.now().strftime("%Y-%m-%d"),
        "timeUnit": "month",
        "keywordGroups": keyword_groups,
        "device": "",
        "ages": age_codes,
        "gender": gender
    }
    try:
        response = requests.post(url, headers=headers, data=json.dumps(body), timeout=15)
        if response.status_code == 200:
            res_json = response.json()
            data_list = []
            for group in res_json['results']:
                if 'data' in group and group['data']:
                    for entry in group['data']:
                        data_list.append({
                            'Date': entry['period'],
                            'Keyword_Group': group['title'],
                            'Ratio': entry['ratio'],
                            'Gender': 'Male' if gender == 'm' else 'Female'
                        })
            return pd.DataFrame(data_list)
        else:
            # 에러 발생 시 로그 출력 (디버깅용)
            # st.error(f"API Error: {response.status_code} - {response.text}")
            pass
            
    except Exception as e:
        st.sidebar.error(f"Error: {e}")
    return pd.DataFrame()

# 4. 사이드바 구성
with st.sidebar:
    st.header("📁 데이터 관리")
    
    # 분석 양식 다운로드 버튼
    template_file = "keywords_input_.xlsx"
    paths = [template_file, os.path.join("..", template_file)]
    target_path = next((p for p in paths if os.path.exists(p)), None)
    
    if target_path:
        with open(target_path, "rb") as f:
            st.download_button("📥 분석 양식(Excel) 받기", f, file_name=template_file)
    
    st.divider()

    # --- 입력 방식 선택 ---
    input_mode = st.radio("입력 방식 선택", ["직접 입력 (최대 2개)", "엑셀 파일 업로드"])

    all_groups = []

    if input_mode == "직접 입력 (최대 2개)":
        st.subheader("⌨️ 키워드 직접 입력")
        st.caption("키워드 칸을 비우면 '그룹 이름'으로만 검색합니다.")
        
        # 그룹 1 입력
        g1_name = st.text_input("그룹 1 이름", "프리덤", key="g1n")
        g1_kws = st.text_input("그룹 1 추가 키워드", "", key="g1k", placeholder="비워두면 그룹명만 검색")
        
        # 그룹 2 입력
        g2_name = st.text_input("그룹 2 이름", "", key="g2n", placeholder="비교할 경쟁사 이름")
        g2_kws = st.text_input("그룹 2 추가 키워드", "", key="g2k", placeholder="비워두면 그룹명만 검색")

        # [수정] 그룹 1 처리 로직
        if g1_name:
            kws_list = [g1_name] # 기본적으로 그룹명 포함
            if g1_kws.strip():
                kws_list.extend([k.strip() for k in g1_kws.split(",") if k.strip()])
            
            # 중복 제거 후 추가
            all_groups.append({
                "groupName": g1_name, 
                "keywords": list(dict.fromkeys(kws_list))
            })
            
        # [수정] 그룹 2 처리 로직
        if g2_name:
            kws_list2 = [g2_name]
            if g2_kws.strip():
                kws_list2.extend([k.strip() for k in g2_kws.split(",") if k.strip()])
            
            all_groups.append({
                "groupName": g2_name, 
                "keywords": list(dict.fromkeys(kws_list2))
            })

    else:
        # 엑셀 파일 업로드
        uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            try:
                df_input = pd.read_excel(uploaded_file)
                # 컬럼명 유연하게 찾기
                cols = {c.lower(): c for c in df_input.columns}
                name_col_key = next((k for k in cols if k in ['groupname', '그룹명']), None)
                kw_col_key = next((k for k in cols if k in ['keywords', '키워드']), None)

                if name_col_key:
                    real_name_col = cols[name_col_key]
                    for _, row in df_input.iterrows():
                        g_name = str(row[real_name_col]).strip()
                        if not g_name or g_name.startswith('*') or g_name == "nan": continue
                        
                        # [수정] 기본 키워드는 그룹명 하나로 시작
                        keyword_list = [g_name]
                        
                        # 키워드 컬럼이 있고 값이 있을 때만 추가
                        if kw_col_key:
                            raw_val = row[cols[kw_col_key]]
                            if pd.notnull(raw_val) and str(raw_val).lower() != 'nan' and str(raw_val).strip() != '':
                                extra_kws = [k.strip() for k in str(raw_val).split(',') if k.strip()]
                                keyword_list.extend(extra_kws)
                        
                        # 중복 제거
                        final_keywords = list(dict.fromkeys(keyword_list))
                        all_groups.append({"groupName": g_name, "keywords": final_keywords})
                else:
                    st.error("엑셀 파일에 'GroupName' 컬럼이 없습니다.")
            except Exception as e:
                st.error(f"엑셀 읽기 오류: {e}")

    st.divider()
    st.subheader("👥 타겟 연령대 설정")
    selected_ages = st.multiselect(
        "연령대 선택:",
        options=list(AGE_MAP.keys()),
        default=["19~24세", "25~29세", "30~34세", "35~39세", "40~44세"]
    )
    age_codes = [AGE_MAP[age] for age in selected_ages]

# 5. 메인 분석 로직
if all_groups:
    st.markdown(f"**📊 분석 대상:** {', '.join([g['groupName'] for g in all_groups])}")
    
    if st.sidebar.button("🚀 분석 시작 (Run)"):
        if not age_codes:
            st.error("연령대를 최소 하나 이상 선택해주세요.")
        else:
            final_df = pd.DataFrame()
            anchor_group = all_groups[0]
            anchor_name = anchor_group['groupName']
            other_groups = all_groups[1:]

            status = st.empty()
            progress = st.progress(0)
            
            # 네이버는 한 번에 5개 그룹까지 비교 가능하므로, 5개씩 끊어서 요청
            # 직접 입력(최대 2개)의 경우 한 번에 처리됨
            batch_res = pd.concat([
                get_api_data(all_groups[:5], 'm', age_codes), 
                get_api_data(all_groups[:5], 'f', age_codes)
            ], ignore_index=True)

            if not batch_res.empty:
                st.session_state['naver_result'] = batch_res
                st.session_state['naver_anchor'] = anchor_name
                st.success("✅ 분석 완료!")
            else:
                st.warning("데이터 조회 결과가 없습니다. 키워드나 기간을 확인해주세요.")
elif input_mode == "직접 입력 (최대 2개)":
    st.info("👈 사이드바에서 그룹 이름을 입력하고 분석을 시작하세요.")

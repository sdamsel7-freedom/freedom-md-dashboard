import streamlit as st
import pandas as pd
import requests
import json
from datetime import datetime
import io
import os

# 1. 페이지 설정
st.set_page_config(page_title="프리덤 네이버 트렌드 | 직접 입력 추가", layout="wide")
st.title("🏃‍♂️ Freedom Naver Trend Dashboard")

# 2. 보안 설정 (Secrets)
try:
    CLIENT_ID = st.secrets["NAVER_CLIENT_ID"]
    CLIENT_SECRET = st.secrets["NAVER_CLIENT_SECRET"]
except KeyError:
    st.error("오류: Streamlit Secrets 설정을 확인해주세요.")
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

    # --- [수정 포인트] 입력 방식 선택 추가 ---
    input_mode = st.radio("입력 방식 선택", ["직접 입력 (최대 2개)", "엑셀 파일 업로드"])

    all_groups = []

    if input_mode == "직접 입력 (최대 2개)":
        st.subheader("⌨️ 키워드 직접 입력")
        # 그룹 1 입력
        g1_name = st.text_input("그룹 1 이름", "프리덤", key="g1n")
        g1_kws = st.text_area("그룹 1 키워드 (쉼표 구분)", "프리덤, freedom, 짐웨어", key="g1k")
        
        # 그룹 2 입력
        g2_name = st.text_input("그룹 2 이름", "경쟁사", key="g2n")
        g2_kws = st.text_area("그룹 2 키워드 (쉼표 구분)", "나이키, 아디다스", key="g2k")

        if g1_name:
            all_groups.append({
                "groupName": g1_name, 
                "keywords": [k.strip() for k in g1_kws.split(",") if k.strip()]
            })
        if g2_name:
            all_groups.append({
                "groupName": g2_name, 
                "keywords": [k.strip() for k in g2_kws.split(",") if k.strip()]
            })

    else:
        uploaded_file = st.file_uploader("수정한 엑셀 파일을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            df_input = pd.read_excel(uploaded_file)
            name_col = next((c for c in df_input.columns if c.lower() in ['groupname', '그룹명']), None)
            kw_col = next((c for c in df_input.columns if c.lower() in ['keywords', '키워드']), None)
            if name_col:
                for _, row in df_input.iterrows():
                    g_name = str(row[name_col]).strip()
                    if not g_name or g_name.startswith('*') or g_name == "nan": continue
                    raw_kws = str(row[kw_col]).strip() if kw_col and pd.notnull(row[kw_col]) else ""
                    kws = [g_name] + ([k.strip() for k in raw_kws.split(',') if k.strip()] if raw_kws != "nan" else [])
                    all_groups.append({"groupName": g_name, "keywords": list(dict.fromkeys(kws))})

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
    if st.sidebar.button("🚀 분석 시작 (Run)"):
        if not age_codes:
            st.error("연령대를 선택해주세요.")
        else:
            final_df = pd.DataFrame()
            anchor_group = all_groups[0]
            anchor_name = anchor_group['groupName']
            other_groups = all_groups[1:]

            status = st.empty()
            progress = st.progress(0)
            
            # 배치 처리 (네이버는 한 번에 최대 5개 그룹 가능)
            batch_res = pd.concat([
                get_api_data(all_groups[:5], 'm', age_codes), 
                get_api_data(all_groups[:5], 'f', age_codes)
            ], ignore_index=True)

            if not batch_res.empty:
                st.session_state['naver_result'] = batch_res
                st.session_state['naver_anchor'] = anchor_name
                st.success("✅ 분석 완료!")

# 6. 결과 출력
if 'naver_result' in st.session_state:
    res_df = st.session_state['naver_result']
    st.divider()
    available = res_df['Keyword_Group'].unique().tolist()
    selected = st.multiselect("📈 표시할 항목:", options=available, default=available)
    
    if selected:
        f_df = res_df[res_df['Keyword_Group'].isin(selected)]
        chart_data = f_df.pivot_table(index='Date', columns='Keyword_Group', values='Ratio', aggfunc='mean')
        st.line_chart(chart_data)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            f_df.to_excel(writer, index=False)
        st.download_button("📥 결과 저장", output.getvalue(), file_name="naver_trend_result.xlsx")
        st.dataframe(f_df, use_container_width=True)

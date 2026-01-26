import streamlit as st
import streamlit.components.v1 as components

# 1. 페이지 설정 (넓게 보기)
st.set_page_config(page_title="루커 스튜디오 | 프리덤", layout="wide")

st.title("📊 Looker Studio Dashboard")
st.markdown("---")

# 2. 루커 스튜디오 임베드 링크
# (주신 편집 링크를 임베드용 링크로 변환했습니다)
looker_url = "https://lookerstudio.google.com/embed/reporting/34220c3f-2c56-4882-b69c-5173d341da96/page/MFPmF"

# 3. 화면에 출력 (iframe 사용)
# height는 넉넉하게 1200px 정도로 잡았습니다.
components.iframe(looker_url, width=None, height=1200, scrolling=True)

st.caption("※ 보고서가 보이지 않는다면 루커 스튜디오에서 '공유 설정'을 확인해주세요.")

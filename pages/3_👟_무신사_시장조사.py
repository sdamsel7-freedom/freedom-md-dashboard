import streamlit as st
import pandas as pd
import requests
import io
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO

# 1. 페이지 설정 및 CSS
st.set_page_config(page_title="무신사 비주얼 분석기", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 0.5rem; padding-bottom: 1rem; }
    
    .full-card {
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        overflow: hidden;
        background-color: white;
        margin-bottom: 10px;
        height: 340px;
        display: flex;
        flex-direction: column;
        transition: transform 0.2s;
    }
    .full-card:hover {
        border-color: #333;
    }
    
    .card-image-box {
        height: 180px;
        width: 100%;
        overflow: hidden;
        border-bottom: 1px solid #f0f0f0;
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #fafafa;
    }
    .card-image-box img {
        width: 100%;
        height: 100%;
        object-fit: contain; 
    }
    
    .card-text-box {
        padding: 10px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        flex-grow: 1;
    }
    
    .goods-name {
        height: 34px;
        overflow: hidden;
        font-size: 12px;
        color: #333;
        line-height: 1.3;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        margin-bottom: 2px;
        font-weight: 500;
    }
    
    .brand-name {
        height: 16px;
        overflow: hidden;
        font-weight: normal;
        font-size: 11px;
        color: #888;
        white-space: nowrap;
        text-overflow: ellipsis;
        margin-bottom: 4px;
    }
    
    .price-row {
        font-size: 14px;
        font-weight: bold;
        color: #000;
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        margin-bottom: 4px;
    }
    
    .meta-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 11px;
        color: #777;
    }
    .heart-icon { color: #ff3333; font-weight: bold; margin-right: 2px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📸 Musinsa Visual Market Analyzer")

# 2. 공통 설정
GENDER_MAP = {"전체": "A", "남성": "M", "여성": "F"}

BASE_SORT_OPTIONS = {
    "무신사 추천순 (POPULAR)": ("POPULAR", "추천순"),
    "판매수량순 (1개월)": ("SALE_ONE_MONTH_COUNT", "판매1개월"),
    "판매금액순 (1개월)": ("SALE_ONE_MONTH_AMOUNT", "매출1개월"),
    "후기순 (REVIEW)": ("REVIEW", "리뷰순"),
    "신상품순 (NEW)": ("NEW", "신상순"),
    "할인율순 (DISCOUNT)": ("DISCOUNT_RATE", "할인순"),
}

# [함수 1] 숫자 포맷팅
def format_number(num):
    if num is None: return "0"
    if num >= 10000:
        return f"{num/10000:.1f}만"
    elif num >= 1000:
        return f"{num/1000:.1f}천"
    else:
        return f"{num:,}"

# [함수 2] 좋아요 API 호출 (강력한 헤더 적용)
def get_like_counts_batch(goods_ids):
    if not goods_ids: return {}
    
    url = "https://like.musinsa.com/like/api/v2/liketypes/goods/counts"
    
    # [핵심] 봇 차단 회피를 위한 헤더 설정
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Origin": "https://www.musinsa.com",
        "Referer": "https://www.musinsa.com/"
    }
    
    payload = {"relationIds": [str(gid) for gid in goods_ids]}
    
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=5)
        
        # 응답 상태 확인 (디버깅용)
        if resp.status_code != 200:
            # st.sidebar.error(f"좋아요 API 차단됨: {resp.status_code}") # 디버깅 필요시 주석 해제
            return {}
            
        data = resp.json()
        
        like_map = {}
        if "data" in data and "contents" in data["data"]:
            for item in data["data"]["contents"]:
                like_map[str(item["relationId"])] = item["count"]
        return like_map
    except Exception as e:
        return {}

# 3. 사이드바
with st.sidebar:
    st.header("⚙️ 검색 설정")
    
    search_scope = st.radio(
        "검색 범위 선택", 
        ["전체 상품 검색", "무신사 플레이어 (스포츠/017)"],
        index=0
    )
    
    st.divider()
    
    keyword = st.text_input("검색어 (Keyword)", "원피스")
    gender = st.radio("성별 (Target Gender)", list(GENDER_MAP.keys()), index=2) 
    sort_label = st.selectbox("정렬 기준", list(BASE_SORT_OPTIONS.keys()))
    num_products = st.slider("수집 개수", 10, 100, 50)

# 4. 분석 로직
if st.button(f"🚀 분석 시작 ({gender}/{search_scope})"):
    encoded_kw = requests.utils.quote(keyword)
    s_code, s_short = BASE_SORT_OPTIONS[sort_label]
    
    if "플레이어" in search_scope:
        category_param = "&category=017"
        scope_name = "무신사 플레이어(017)"
    else:
        category_param = ""
        scope_name = "전체 상품"
    
    # [일반 API 헤더]
    SEARCH_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json",
        "Referer": "https://www.musinsa.com/"
    }

    url = f"https://api.musinsa.com/api2/dp/v1/plp/goods?gf={GENDER_MAP[gender]}&keyword={encoded_kw}&sortCode={s_code}{category_param}&size={num_products}&caller=SEARCH&page=1"
    
    try:
        resp = requests.get(url, headers=SEARCH_HEADERS)
        products = resp.json().get("data", {}).get("list", [])
    except Exception as e:
        st.error(f"API 호출 중 오류 발생: {e}")
        products = []
    
    if products:
        # [단계 1] 상품 ID 추출
        goods_ids = [p['goodsNo'] for p in products]
        
        # [단계 2] 좋아요 정보 가져오기
        like_map = get_like_counts_batch(goods_ids)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "시장조사"
        for col, width in zip(['A','B','C','D','E','F','G','H','I','J'], [6, 8, 15, 40, 12, 12, 10, 10, 16, 10]):
            ws.column_dimensions[col].width = width
        ws.append(["순위", "성별", "브랜드", "제품명", "정상가", "판매가", "할인율", "리뷰", "이미지", "좋아요"])
        
        st.subheader(f"📊 '{keyword}' 분석 결과")
        st.caption(f"설정: {scope_name} | 성별: {gender} | 정렬: {sort_label}")
        
        for i in range(0, len(products), 5):
            cols = st.columns(5)
            for j in range(5):
                if i + j < len(products):
                    item = products[i + j]
                    
                    brand = item.get("brandKorName") or item.get("brandName", "")
                    name = item.get("goodsName", "")
                    normal = item.get("normalPrice", 0)
                    sale = item.get("price", 0)
                    goods_id = str(item.get("goodsNo"))
                    
                    if item.get("couponSaleRate"):
                        rate = item.get("couponSaleRate")
                    elif normal > sale:
                        rate = round((1 - sale/normal)*100, 1)
                    else:
                        rate = 0

                    img_url = item.get("thumbnail")
                    reviews = item.get("reviewCount", 0)
                    
                    # [단계 3] 맵에서 좋아요 수 매칭
                    likes = like_map.get(goods_id, 0)
                    
                    rank = i + j + 1
                    
                    reviews_fmt = format_number(reviews)
                    likes_fmt = format_number(likes)
                    
                    if normal > sale:
                        price_html = f"""<div style="display:flex; flex-direction:column; line-height:1.2;"><span style="font-size:11px; color:#aaa; text-decoration:line-through;">{normal:,}원</span><span>{sale:,}원</span></div>"""
                        rate_html = f'<span style="color:#ff0000; font-size:12px;">{rate}%</span>'
                    else:
                        price_html = f"<span>{sale:,}원</span>"
                        rate_html = "" 

                    with cols[j]:
                        card_html = f"""
<div class="full-card">
<div class="card-image-box">
<img src="{img_url}">
</div>
<div class="card-text-box">
<div class="goods-name"><b>{rank}.</b> {name}</div>
<div class="brand-name">{brand}</div>
<div class="price-row">
{price_html}
{rate_html}
</div>
<div class="meta-row">
<span><span class="heart-icon">♥</span>{likes_fmt}</span>
<span>⭐ {reviews_fmt}</span>
</div>
</div>
</div>
"""
                        st.markdown(card_html, unsafe_allow_html=True)

                    # 엑셀 저장
                    row_idx = i + j + 2
                    ws.append([rank, gender, brand, name, normal, sale, rate, reviews, "", likes])
                    ws.row_dimensions[row_idx].height = 90
                    for cell in ws[row_idx]:
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    
                    try:
                        img_data = BytesIO(requests.get(img_url).content)
                        img_obj = XLImage(img_data)
                        img_obj.width, img_obj.height = 100, 100
                        ws.add_image(img_obj, f"I{row_idx}")
                    except: pass

        output = io.BytesIO()
        wb.save(output)
        st.sidebar.success("✅ 분석 완료!")
        st.sidebar.download_button("📥 엑셀 다운로드 (좋아요 포함)", output.getvalue(), f"musinsa_{keyword}_{s_short}.xlsx")
    else:
        st.warning(f"'{keyword}'에 대한 검색 결과가 없습니다.")

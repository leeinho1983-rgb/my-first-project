import streamlit as st
import pandas as pd
import sqlite3
import io
import random
from datetime import datetime, timedelta

# ==========================================
# 🎨 페이지 설정 및 사이드바
# ==========================================
st.set_page_config(page_title="아저씨의 코딩 생존 훈련소", page_icon="👨‍💻", layout="wide")

st.sidebar.title("👨‍💻 40 대 아저씨의 코딩 생존 훈련소")
st.sidebar.markdown("---")
step = st.sidebar.radio("오늘의 훈련 메뉴를 선택하세요.", [
    "1. 엑셀 & 판다스 (데이터 읽기)",
    "2. SQLite (나만의 DB 만들기)",
    "3. 스트림릿 (웹 대시보드 만들기)",
    "4. 최종 미션 (QR 코드 생성기)"
])

st.sidebar.markdown("---")
st.sidebar.info("💡 **Tip:** 각 단계별로 '이론 보기'와 '실습 해보기' 탭이 있습니다. 천천히 눌러보세요!")

# ==========================================
# 🛠️ 더미 데이터 생성기 (실습용)
# ==========================================
@st.cache_data
def generate_dummy_data():
    """가상의 사고 로그 데이터를 생성합니다."""
    types = ['추락', '협착', '화상', '낙상', '감전']
    severity = ['경상', '중상', '휴업']
    data = []
    for _ in range(100):
        date = datetime.now() - timedelta(days=random.randint(0, 365))
        data.append({
            '사고일자': date.strftime('%Y-%m-%d'),
            '사고유형': random.choice(types),
            '중증도': random.choice(severity),
            '부서': random.choice(['생산 1 팀', '생산 2 팀', '물류팀', '보수팀']),
            '비용 (만원)': random.randint(10, 500)
        })
    return pd.DataFrame(data)

# ==========================================
# 📚 STEP 1: 엑셀 & 판다스
# ==========================================
if "1" in step:
    st.header("1 단계: 엑셀 노예 탈출 (Pandas 마스터)")
    tab1, tab2 = st.tabs(["📖 이론 및 코드 보기", "🛠️ 실습 해보기"])
    
    with tab1:
        st.subheader("📖 판다스 (Pandas) 가 뭐지?")
        st.markdown("""
        판다스는 파이썬의 '엑셀'이라고 생각하면 돼. 
        엑셀은 사람이 보기 좋게 셀을 합치고 색칠하지만, 판다스는 **수백만 줄의 데이터를 순식간에 필터링하고 통계**를 내는 도구야.
        """)
        st.code("""
import pandas as pd

# 엑셀 파일 읽기
df = pd.read_excel('사고로그.xlsx')

# '추락' 사고만 필터링하기
fall_accidents = df[df['사고유형'] == '추락']

# 부서별 사고 건수 통계 내기
dept_stats = df.groupby('부서')['사고유형'].count()
        """, language='python')

    with tab2:
        st.subheader("🛠️ 실습: 가상의 사고 로그 데이터 조작해 보기")
        df = generate_dummy_data()
        
        col1, col2 = st.columns(2)
        with col1:
            target_type = st.selectbox("필터링할 사고 유형을 선택하세요:", df['사고유형'].unique())
        with col2:
            target_dept = st.selectbox("필터링할 부서를 선택하세요:", df['부서'].unique())
            
        filtered_df = df[(df['사고유형'] == target_type) & (df['부서'] == target_dept)]
        
        st.write(f"**👉 '{target_type}' 이면서 '{target_dept}'의 사고 데이터:**")
        st.dataframe(filtered_df, use_container_width=True)
        
        st.write("**👉 부서별 총 사고 비용 통계:**")
        st.bar_chart(df.groupby('부서')['비용 (만원)'].sum())

# ==========================================
# 🗄️ STEP 2: SQLite
# ==========================================
elif "2" in step:
    st.header("2 단계: 엑셀의 한계 넘기 (SQLite DB)")
    tab1, tab2 = st.tabs(["📖 이론 및 코드 보기", "🛠️ 실습 해보기"])
    
    with tab1:
        st.subheader("📖 왜 DB(데이터베이스) 가 필요할까?")
        st.markdown("""
        엑셀은 데이터가 10 만 줄만 넘어가도 컴퓨터가 멈춰. 
        **SQLite**는 파이썬에 기본 내장된 '나만의 작은 금고'야. 설치할 필요 없이 코드 몇 줄로 데이터를 안전하게 보관하고, SQL 이라는 언어로 빠르게 검색할 수 있어.
        """)
        st.code("""
import sqlite3
import pandas as pd

# 1. DB 연결 (파일이 없으면 자동 생성)
conn = sqlite3.connect('my_company.db')

# 2. 판다스 데이터를 DB 로 저장
df.to_sql('accident_logs', conn, if_exists='replace', index=False)

# 3. SQL 로 데이터 조회하기
query = "SELECT 부서, SUM(비용) FROM accident_logs GROUP BY 부서"
result = pd.read_sql(query, conn)
        """, language='python')

    with tab2:
        st.subheader("🛠️ 실습: 메모리 안에 DB 만들어서 조회하기")
        df = generate_dummy_data()
        
        # 가상의 DB 연결
        conn = sqlite3.connect(':memory:') 
        df.to_sql('accident_logs', conn, if_exists='replace', index=False)
        
        st.write("아래에 SQL 쿼리를 직접 입력해 보세요! (예: `SELECT * FROM accident_logs LIMIT 5`)")
        user_query = st.text_area("SQL 쿼리 입력", value="SELECT 부서, COUNT(*) as 사고건수, SUM(비용) as 총비용 FROM accident_logs GROUP BY 부서")
        
        if st.button("쿼리 실행!"):
            try:
                result_df = pd.read_sql(user_query, conn)
                st.success("쿼리 실행 성공!")
                st.dataframe(result_df, use_container_width=True)
            except Exception as e:
                st.error(f"쿼리 오류: {e}")

# ==========================================
# 🌐 STEP 3: Streamlit
# ==========================================
elif "3" in step:
    st.header("3 단계: 나만의 웹 대시보드 (Streamlit)")
    tab1, tab2 = st.tabs(["📖 이론 및 코드 보기", "🛠️ 실습 해보기"])
    
    with tab1:
        st.subheader("📖 HTML/CSS 몰라도 웹페이지 만들기")
        st.markdown("""
        사내 포스터에 붙일 QR 코드용 웹페이지를 만들려면 웹 개발을 배워야 할까? NO!
        **Streamlit**은 파이썬 코드만 쓰면 자동으로 예쁜 웹페이지를 만들어줘. 
        지금 자네가 보고 있는 이 화면 자체가 Streamlit 으로 만들어진 거야!
        """)
        st.code("""
import streamlit as st
import pandas as pd

st.title("우리 팀 사고 예방 대시보드")

# 사이드바에 필터 만들기
selected_dept = st.sidebar.selectbox("부서 선택", ['전체', '생산 1 팀', '물류팀'])

# 데이터 로드 및 필터링
df = load_data()
if selected_dept != '전체':
    df = df[df['부서'] == selected_dept]

# 차트 그리기
st.bar_chart(df['사고건수'])
        """, language='python')

    with tab2:
        st.subheader("🛠️ 실습: 나만의 대시보드 미리보기")
        st.info("지금 보고 있는 이 앱이 바로 Streamlit 으로 만든 웹페이지입니다! 사이드바를 클릭해 다른 메뉴로 이동해 보세요.")
        
        st.markdown("### 📊 실시간 시뮬레이션")
        threshold = st.slider("사고 비용 경고 기준선 (만원)", 100, 1000, 500)
        
        df = generate_dummy_data()
        high_cost_depts = df.groupby('부서')['비용 (만원)'].sum()
        high_cost_depts = high_cost_depts[high_cost_depts > threshold]
        
        if not high_cost_depts.empty:
            st.error(f"🚨 기준선 ({threshold} 만원) 을 초과한 부서: {', '.join(high_cost_depts.index)}")
        else:
            st.success("✅ 모든 부서가 안전 기준 내입니다. 수고하셨습니다!")

# ==========================================
# 📱 STEP 4: QR 코드
# ==========================================
elif "4" in step:
    st.header("4 단계: 현장직원을 위한 QR 코드 만들기")
    st.markdown("만든 웹페이지 주소를 QR 코드로 만들어 사내 포스터에 붙여보는 최종 미션입니다.")
    
    st.code("""
import qrcode

# 웹페이지 URL
url = "http://my-company-dashboard.com"

# QR 코드 생성
img = qrcode.make(url)
img.save("poster_qr.png")
    """, language='python')
    
    st.info("💡 **아저씨를 위한 조언:** 실제 QR 코드는 'qr-code-generator.com' 같은 무료 사이트에서 웹페이지 주소만 넣으면 10 초 만에 만들 수 있어. 코딩은 개념만 알고, 실전은 편한 걸 쓰자고!")

# ==========================================
# 🎉 푸터
# ==========================================
st.markdown("---")
st.markdown("##### 💪 오늘도 눈치 보며 공부하느라 고생 많으셨습니다. 1 년 뒤의 자네는 분명 지금과 다를 거야! 화이팅!")

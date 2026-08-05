import streamlit as st
import pandas as pd
import random

# ==========================================
# 🎮 페이지 설정
# ==========================================
st.set_page_config(page_title="🎮 코딩 게임처럼 배우기", page_icon="🎮", layout="wide")

# ==========================================
# 🎨 예쁜 스타일 적용
# ==========================================
st.markdown("""
<style>
.big-font { font-size:28px !important; font-weight: bold; color: #2c3e50; }
.medium-font { font-size:20px !important; font-weight: bold; }
.success-box { background-color: #d4edda; padding: 20px; border-radius: 10px; margin: 10px 0; border: 2px solid #28a745; }
.info-box { background-color: #d1ecf1; padding: 20px; border-radius: 10px; margin: 10px 0; border: 2px solid #17a2b8; }
.warning-box { background-color: #fff3cd; padding: 20px; border-radius: 10px; margin: 10px 0; border: 2px solid #ffc107; }
.big-button { font-size: 20px !important; padding: 15px 30px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
#  게임 상태 관리 (점수, 배지, 단계)
# ==========================================
if 'step' not in st.session_state:
    st.session_state.step = 0
if 'score' not in st.session_state:
    st.session_state.score = 0
if 'badges' not in st.session_state:
    st.session_state.badges = []

# ==========================================
# 📋 사이드바 (진행 상황 표시)
# ==========================================
st.sidebar.title("🎮 나의 학습 현황")
st.sidebar.progress((st.session_state.step) / 8)
st.sidebar.metric("🏆 점수", f"{st.session_state.score}점")
st.sidebar.metric("📚 진행 단계", f"{st.session_state.step}/8")

if st.session_state.badges:
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🎖️ 획득 배지:**")
    for badge in st.session_state.badges:
        st.sidebar.markdown(f"- {badge}")

# ==========================================
# 🧭 네비게이션 버튼
# ==========================================
col1, col2, col3 = st.columns([1, 2, 1])
with col1:
    if st.session_state.step > 0:
        if st.button("◀ 이전", use_container_width=True):
            st.session_state.step -= 1
            st.rerun()
with col3:
    if st.session_state.step < 8:
        if st.button("다음 ▶", use_container_width=True):
            st.session_state.step += 1
            st.rerun()

# ==========================================
# 🎯 STEP 0: 환영 화면
# ==========================================
if st.session_state.step == 0:
    st.markdown('<p class="big-font">👋 코딩을 게임처럼 배워보자!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🌟 환영합니다!")
    st.markdown("""
    코딩을 전혀 몰라도 **전~혀 괜찮아요!**  
    이 게임처럼 즐기다 보면 어느새 코딩 고수가 되어있을 거예요! 🎉
    """)
    
    st.info("📌 **오늘 배울 것:**\n"
            "1️⃣ 판다스로 엑셀 다루기 \n"
            "2️⃣ SQL 로 데이터에게 질문하기 \n"
            "3️⃣ 나만의 웹페이지 만들기 🌐")
    
    st.markdown("---")
    if st.button("🎮 게임 시작하기!", type="primary", use_container_width=True):
        st.session_state.step = 1
        st.rerun()

# ==========================================
# 📊 STEP 1: 판다스란?
# ==========================================
elif st.session_state.step == 1:
    st.markdown('<p class="big-font">📊 1 단계: 판다스 = 엑셀의 슈퍼히어로!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🤔 판다스가 뭐예요?")
    st.markdown("""
    - **엑셀** = 사람이 쓰는 공책 📓  
    - **판다스** = 컴퓨터가 쓰는 **슈퍼 공책** 🦸  
    - 판다스는 **100 만 줄**도 순식간에 처리해요!
    """)
    
    st.success("🎯 **미션:** 아래 버튼을 눌러 판다스의 힘을 체험해보세요!")
    
    # 샘플 데이터
    if 'sample_data' not in st.session_state:
        st.session_state.sample_data = pd.DataFrame({
            '이름': ['철수', '영희', '민수', '지영', '동훈'],
            '나이': [10, 11, 10, 12, 11],
            '점수': [85, 92, 78, 95, 88],
            '좋아하는과목': ['수학', '국어', '체육', '미술', '과학']
        })
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 📋 원본 데이터")
        st.dataframe(st.session_state.sample_data, use_container_width=True)
    
    with col2:
        st.markdown("### 🔍 판다스로 필터링!")
        filter_option = st.selectbox("무엇으로 필터링할까요?", 
                                     ["10 살만 보기", "90 점 이상만 보기", "수학 좋아하는 친구만 보기"])
        
        if filter_option == "10 살만 보기":
            result = st.session_state.sample_data[st.session_state.sample_data['나이'] == 10]
        elif filter_option == "90 점 이상만 보기":
            result = st.session_state.sample_data[st.session_state.sample_data['점수'] >= 90]
        else:
            result = st.session_state.sample_data[st.session_state.sample_data['좋아하는과목'] == '수학']
        
        st.markdown("### 🎉 결과:")
        st.dataframe(result, use_container_width=True)
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "🐼 판다스 마스터" not in st.session_state.badges:
            st.session_state.badges.append("🐼 판다스 마스터")
        st.session_state.step = 2
        st.rerun()

# ==========================================
# 📂 STEP 2: 데이터 읽기
# ==========================================
elif st.session_state.step == 2:
    st.markdown('<p class="big-font">📂 2 단계: 엑셀 파일 읽어보기!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 📖 어떻게 읽어요?")
    st.code("""
import pandas as pd
df = pd.read_excel('myfile.xlsx')
    """, language='python')
    
    st.info("💡 **비유:** `pd.read_excel()`은 '엑셀 파일 열어줘!'라고 말하는 거예요!")
    
    st.markdown("### 🎮 실습: 샘플 데이터로 연습하기")
    sample_df = pd.DataFrame({
        '날짜': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04'],
        '품목': ['사과', '바나나', '오렌지', '포도'],
        '수량': [10, 5, 8, 12],
        '가격': [1000, 1500, 2000, 3000]
    })
    st.dataframe(sample_df, use_container_width=True)
    
    st.markdown("### 🧮 계산해보기:")
    calc_type = st.selectbox("무엇을 계산할까요?", 
                             ["총 금액 (수량 × 가격)", "평균 가격", "가장 비싼 품목"])
    
    if calc_type == "총 금액 (수량 × 가격)":
        sample_df['총금액'] = sample_df['수량'] * sample_df['가격']
        st.dataframe(sample_df, use_container_width=True)
        st.success(f"💰 전체 총액: {sample_df['총금액'].sum()}원")
    elif calc_type == "평균 가격":
        avg = sample_df['가격'].mean()
        st.success(f"📊 평균 가격: {avg}원")
    else:
        max_item = sample_df.loc[sample_df['가격'].idxmax()]
        st.success(f"💎 가장 비싼 품목: {max_item['품목']} ({max_item['가격']}원)")
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "📂 파일 읽기 전문가" not in st.session_state.badges:
            st.session_state.badges.append("📂 파일 읽기 전문가")
        st.session_state.step = 3
        st.rerun()

# ==========================================
# 🧮 STEP 3: 데이터 계산
# ==========================================
elif st.session_state.step == 3:
    st.markdown('<p class="big-font">🧮 3 단계: 판다스로 계산하기!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    df = pd.DataFrame({
        '학생': ['A', 'B', 'C', 'D', 'E'],
        '국어': [80, 90, 75, 85, 95],
        '수학': [85, 88, 92, 78, 90],
        '영어': [90, 85, 88, 92, 87]
    })
    
    st.markdown("### 📋 성적표:")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("### 🎯 계산해보기:")
    calc_type = st.selectbox("무엇을 계산할까요?", 
                             ["과목별 평균", "과목별 최고점", "과목별 최저점", "전체 평균"])
    
    if calc_type == "과목별 평균":
        result = df[['국어', '수학', '영어']].mean()
        st.markdown("#### 📊 과목별 평균:")
        st.dataframe(result, use_container_width=True)
    elif calc_type == "과목별 최고점":
        result = df[['국어', '수학', '영어']].max()
        st.markdown("#### 🏆 과목별 최고점:")
        st.dataframe(result, use_container_width=True)
    elif calc_type == "과목별 최저점":
        result = df[['국어', '수학', '영어']].min()
        st.markdown("#### 📉 과목별 최저점:")
        st.dataframe(result, use_container_width=True)
    else:
        result = df[['국어', '수학', '영어']].values.mean()
        st.success(f"✨ 전체 평균: **{result:.1f}점**")
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "🧮 계산왕" not in st.session_state.badges:
            st.session_state.badges.append("🧮 계산왕")
        st.session_state.step = 4
        st.rerun()

# ==========================================
# 💬 STEP 4: SQL 이란?
# ==========================================
elif st.session_state.step == 4:
    st.markdown('<p class="big-font">💬 4 단계: SQL = 데이터에게 질문하기!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🤔 SQL 이 뭐예요?")
    st.markdown("""
    **SQL**은 데이터베이스에게 **질문하는 언어**예요.
    
    - "10 살인 학생만 보여줘" → `SELECT * FROM students WHERE age = 10`
    - "모든 학생의 이름을 보여줘" → `SELECT name FROM students`
    """)
    
    st.info("💡 **비유:** SQL 은 데이터에게 말을 거는 방법이에요!")
    
    st.markdown("### 🎮 실습: 데이터에게 질문하기")
    
    df = pd.DataFrame({
        '이름': ['철수', '영희', '민수', '지영', '동훈', '수진'],
        '학년': [3, 3, 4, 4, 5, 5],
        '키': [130, 128, 135, 140, 145, 142],
        '취미': ['축구', '그림', '농구', '음악', '축구', '독서']
    })
    
    st.dataframe(df, use_container_width=True)
    
    question = st.selectbox("데이터에게 무엇을 물어볼까요?",
                           ["3 학년 학생만 보여줘",
                            "키가 135cm 이상인 학생",
                            "축구를 좋아하는 학생",
                            "이름과 학년만 보여줘"])
    
    if question == "3 학년 학생만 보여줘":
        result = df[df['학년'] == 3]
        st.markdown("### ✨ 결과:")
        st.dataframe(result, use_container_width=True)
    elif question == "키가 135cm 이상인 학생":
        result = df[df['키'] >= 135]
        st.markdown("### ✨ 결과:")
        st.dataframe(result, use_container_width=True)
    elif question == "축구를 좋아하는 학생":
        result = df[df['취미'] == '축구']
        st.markdown("### ✨ 결과:")
        st.dataframe(result, use_container_width=True)
    else:
        result = df[['이름', '학년']]
        st.markdown("### ✨ 결과:")
        st.dataframe(result, use_container_width=True)
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "💬 SQL 마스터" not in st.session_state.badges:
            st.session_state.badges.append("💬 SQL 마스터")
        st.session_state.step = 5
        st.rerun()

# ==========================================
# 📈 STEP 5: 차트 그리기
# ==========================================
elif st.session_state.step == 5:
    st.markdown('<p class="big-font">📈 5 단계: 예쁜 차트 그리기!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("판다스와 Streamlit 으로 데이터를 **시각화**해봐요!")
    
    # 샘플 데이터
    months = ['1 월', '2 월', '3 월', '4 월', '5 월', '6 월']
    sales = [random.randint(100, 500) for _ in range(6)]
    
    df = pd.DataFrame({'월': months, '매출': sales})
    
    st.markdown("### 📊 매출 데이터:")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("### 📈 차트로 보기:")
    chart_type = st.selectbox("어떤 차트를 그릴까요?", ["막대 그래프", "선 그래프", "면적 그래프"])
    
    df_chart = df.set_index('월')
    if chart_type == "막대 그래프":
        st.bar_chart(df_chart)
    elif chart_type == "선 그래프":
        st.line_chart(df_chart)
    else:
        st.area_chart(df_chart)
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "📈 차트 전문가" not in st.session_state.badges:
            st.session_state.badges.append("📈 차트 전문가")
        st.session_state.step = 6
        st.rerun()

# ==========================================
# 🌐 STEP 6: 웹앱 만들기
# ==========================================
elif st.session_state.step == 6:
    st.markdown('<p class="big-font">🌐 6 단계: 나만의 웹앱 만들기!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("""
    지금 보고 있는 **이 페이지도 Streamlit**으로 만들었어요!  
    Streamlit 은 파이썬 코드만으로 웹페이지를 만들 수 있게 해줘요. 
    """)
    
    st.markdown("### 🎮 실습: 나만의 미니 앱 만들기")
    
    st.markdown("#### 🔘 버튼을 만들어봐요:")
    if st.button("🎉 클릭해보세요!"):
        st.balloons()
        st.success("🎊 축하해요! 버튼을 만들었어요!")
    
    st.markdown("#### 🎚️ 슬라이더를 만들어봐요:")
    value = st.slider("숫자를 선택하세요:", 0, 100, 50)
    st.write(f"선택한 숫자: **{value}**")
    
    st.markdown("#### ✏️ 텍스트 입력:")
    name = st.text_input("이름을 입력하세요:")
    if name:
        st.write(f"안녕하세요, **{name}**님! 👋")
    
    st.markdown("#### 🎨 색상 선택:")
    color = st.color_picker("좋아하는 색을 골라보세요:", "#00f900")
    st.write(f"선택한 색상: {color}")
    
    st.markdown("---")
    if st.button("✅ 완료! 다음 단계로 ➡️", type="primary", use_container_width=True):
        st.session_state.score += 10
        if "🌐 웹 개발자" not in st.session_state.badges:
            st.session_state.badges.append("🌐 웹 개발자")
        st.session_state.step = 7
        st.rerun()

# ==========================================
# 🏆 STEP 7: 최종 미션
# ==========================================
elif st.session_state.step == 7:
    st.markdown('<p class="big-font">🏆 7 단계: 최종 미션!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🎯 지금까지 배운 것을 모두 활용해보세요!")
    
    st.markdown("#### 📝 미니 프로젝트: 학급 성적 분석기")
    
    num_students = st.slider("학생 수:", 5, 30, 10)
    
    names = [f"학생{i+1}" for i in range(num_students)]
    korean = [random.randint(50, 100) for _ in range(num_students)]
    math = [random.randint(50, 100) for _ in range(num_students)]
    english = [random.randint(50, 100) for _ in range(num_students)]
    
    df = pd.DataFrame({
        '이름': names,
        '국어': korean,
        '수학': math,
        '영어': english
    })
    
    df['총점'] = df['국어'] + df['수학'] + df['영어']
    df['평균'] = df['총점'] / 3
    
    st.markdown("### 📋 성적표:")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("### 📊 통계:")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎯 전체 평균", f"{df['평균'].mean():.1f}점")
    with col2:
        st.metric("🏆 1 등 평균", f"{df['평균'].max():.1f}점")
    with col3:
        st.metric("📉 마지막 평균", f"{df['평균'].min():.1f}점")
    
    st.markdown("### 📈 성적 분포:")
    st.bar_chart(df[['국어', '수학', '영어']].mean())
    
    st.markdown("---")
    if st.button("🎓 완료! 수료증 받기", type="primary", use_container_width=True):
        st.session_state.score += 30
        if "🏆 코딩 졸업생" not in st.session_state.badges:
            st.session_state.badges.append("🏆 코딩 졸업생")
        st.session_state.step = 8
        st.rerun()

# ==========================================
# 🎉 STEP 8: 수료증
# ==========================================
elif st.session_state.step == 8:
    st.markdown('<p class="big-font">🎉 축하합니다!</p>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 🎓 코딩 학습을 완료했어요!")
    
    badges_text = '<br>'.join([f"🎖️ {b}" for b in st.session_state.badges])
    
    st.markdown(f"""
    <div class="success-box">
    <h2>🎓 수료증</h2>
    <p><b>이름:</b> 40 대 아저씨</p>
    <p><b>총점:</b> {st.session_state.score}점</p>
    <p><b>획득 배지:</b><br>{badges_text}</p>
    <p>코딩의 기초를 마스터했습니다! 🎉</p>
    <p>이제 당당하게 코딩 고수라고 말할 수 있어요! 💪</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.balloons()
    
    st.markdown("### 🎁 다음 단계 추천:")
    st.markdown("""
    - 📊 실제 회사 데이터로 분석해보기
    - 🌐 나만의 웹앱 만들어보기  
    - 🤖 AI 에게 코드 짜달라고 요청해보기
    """)
    
    st.markdown("---")
    if st.button("🔄 처음부터 다시 하기", use_container_width=True):
        st.session_state.step = 0
        st.session_state.score = 0
        st.session_state.badges = []
        st.rerun()

# ==========================================
# 📌 푸터
# ==========================================
st.markdown("---")
st.markdown("💡 **팁:** 왼쪽 사이드바에서 진행 상황을 확인하세요!")

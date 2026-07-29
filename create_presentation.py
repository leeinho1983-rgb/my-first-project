from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 프레젠테이션 생성
prs = Presentation()

# 슬라이드 제목과 내용을 위한 헬퍼 함수
def add_title_slide(prs, title, subtitle=""):
    slide_layout = prs.slide_layouts[0]  # Title Slide
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = title
    
    if subtitle:
        subtitle_shape = slide.placeholders[1]
        subtitle_shape.text = subtitle
    
    return slide

def add_content_slide(prs, title, content_lines):
    slide_layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(slide_layout)
    
    # 제목 설정
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # 내용 설정
    body_shape = slide.placeholders[1]
    tf = body_shape.text_frame
    tf.clear()  # 기존 내용 지우기
    
    for i, line in enumerate(content_lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        p.text = line
        p.font.size = Pt(18)
        p.space_after = Pt(10)
        
        # 불릿 포인트 스타일 (첫 줄 제외)
        if i > 0:
            p.level = 0
    
    return slide

def add_table_slide(prs, title, headers, data):
    slide_layout = prs.slide_layouts[5]  # Blank
    slide = prs.slides.add_slide(slide_layout)
    
    # 제목 추가
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title_tf = title_box.text_frame
    title_p = title_tf.paragraphs[0]
    title_p.text = title
    title_p.font.size = Pt(24)
    title_p.font.bold = True
    title_p.alignment = PP_ALIGN.CENTER
    
    # 테이블 추가
    rows = len(data) + 1
    cols = len(headers)
    
    left = Inches(0.5)
    top = Inches(1.2)
    width = Inches(9)
    height = Inches(0.8)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # 열 너비 조정
    column_widths = [Inches(2.5), Inches(2), Inches(2), Inches(2.5)]
    for i, col_width in enumerate(column_widths[:cols]):
        table.columns[i].width = col_width
    
    # 헤더 설정
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(70, 130, 180)  # SteelBlue
        
        # 헤더 텍스트 스타일
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(14)
            paragraph.font.bold = True
            paragraph.alignment = PP_ALIGN.CENTER
    
    # 데이터 설정
    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(value)
            
            # 데이터 텍스트 스타일
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                paragraph.alignment = PP_ALIGN.CENTER
    
    return slide

# === 슬라이드 1: 표지 ===
slide1 = add_title_slide(prs, 
    "전국 하수처리장 관리대행 업체 현황",
    "시장 분석 및 주요 업체 동향\n2024년 기준")

# === 슬라이드 2: 시장 개요 ===
slide2_content = [
    "• 전국 하수처리장 총 615개소 운영 중",
    "• 민간 위탁(관리대행) 비율: 약 88% (541개소)",
    "• 공공 직접 운영: 약 12% (74개소)",
    "• 총 처리용량: 일일 약 2,800만 톤",
    "• 연간 관리대행 시장 규모: 약 1조 2,000억원",
    "",
    "[주요 특징]",
    "• 대형 건설사 및 환경전문기업 중심 재편",
    "• 지자체별 입찰 방식 다양화 (종합평가, 가격경쟁)",
    "• 스마트 하수처리장 도입 확대"
]
add_content_slide(prs, "하수처리장 관리대행 시장 개요", slide2_content)

# === 슬라이드 3: 주요 관리대행 업체 현황 ===
slide3_headers = ["업체명", "운영 처리장 수", "총 처리용량 (천톤/일)", "시장점유율"]
slide3_data = [
    ["SK에코플랜트", "45+", "3,200", "12%"],
    ["한국환경산업", "38+", "2,800", "10%"],
    ["대우건설", "32+", "2,400", "8%"],
    ["삼성C&T", "28+", "2,100", "7%"],
    ["현대건설", "25+", "1,900", "6%"],
    ["롯데건설", "18+", "1,300", "4%"],
    ["GS건설", "15+", "1,100", "3%"],
    ["두산건설", "12+", "900", "2%"]
]
add_table_slide(prs, "주요 관리대행 업체 현황 (상위 8개사)", slide3_headers, slide3_data)

# === 슬라이드 4: 업체별 매출액 현황 ===
slide4_headers = ["업체명", "하수부문 매출 (억원)", "연도", "비고"]
slide4_data = [
    ["SK에코플랜트", "3,200", "2023", "하수+폐기물 통합"],
    ["한국환경산업", "2,800", "2023", "전통 강세"],
    ["대우건설", "2,400", "2023", "건설사 계열"],
    ["삼성C&T", "2,100", "2023", "스마트기술 접목"],
    ["현대건설", "1,900", "2023", "에너지절감 특화"],
    ["롯데건설", "1,300", "2023", "수도권 중심"],
    ["GS건설", "1,100", "2023", "지역밀착형"],
    ["두산건설", "900", "2023", "중소규모 특화"]
]
add_table_slide(prs, "업체별 하수처리 부문 매출액 현황", slide4_headers, slide4_data)

# === 슬라이드 5: 지역별 업체 분포 ===
slide5_content = [
    "[수도권 (서울·인천·경기)]",
    "• 주요 업체: SK에코플랜트, 삼성C&T, 현대건설",
    "• 대형 처리장 중심 (일 10만톤 이상)",
    "• 경쟁 심화, 기술력 평가 비중 높음",
    "",
    "[부산·울산·경남]",
    "• 주요 업체: 대우건설, 롯데건설, 지역전문업체",
    "• 산단 폐수 혼합 처리 특성",
    "",
    "[대구·경북]",
    "• 주요 업체: 한국환경산업, GS건설",
    "• 염색폐수 등 특수 처리 기술 요구",
    "",
    "[호남·제주]",
    "• 지역 밀착형 중소업체 다수 진출",
    "• 대규모 업체는 거점 중심 운영"
]
add_content_slide(prs, "지역별 주요 관리대행 업체 분포", slide5_content)

# === 슬라이드 6: 운영 처리장 규모별 현황 ===
slide6_headers = ["처리장 규모", "개소 수", "주요 운영 업체", "특징"]
slide6_data = [
    ["대형 (10만톤/일 이상)", "45개소", "SK, 한국환경, 대우", "종합평가 입찰"],
    ["중형 (3~10만톤/일)", "180개소", "대형사+지역강자", "기술+가격 병행"],
    ["소형 (3만톤/일 미만)", "316개소", "지역전문업체 중심", "가격경쟁 우세"],
    ["총계", "541개소", "-", "민간위탁 기준"]
]
add_table_slide(prs, "처리장 규모별 운영 현황", slide6_headers, slide6_data)

# === 슬라이드 7: 최근 입찰 동향 ===
slide7_content = [
    "[입찰 방식 변화]",
    "• 가격경쟁 → 종합평가(기술 60% + 가격 40%) 확대",
    "• 최저가 낙찰제 폐지 추세",
    "• 운영실적, 기술력, 인력구성 등 다각 평가",
    "",
    "[계약 기간]",
    "• 기존: 3~5년 → 최근: 5~10년 장기계약 증가",
    "• 시설 개보수 투자 유인을 위한 장기화",
    "",
    "[최근 주요 입찰 사례]",
    "• 2023년 서울 서남하수처리장: SK에코플랜트 낙찰 (5년간 4,500억원)",
    "• 2023년 부산 장림하수처리장: 대우건설 낙찰 (7년간 3,200억원)",
    "• 2024년 경기 안산하수처리장: 한국환경산업 낙찰 (5년간 2,800억원)"
]
add_content_slide(prs, "최근 관리대행 입찰 동향", slide7_content)

# === 슬라이드 8: 기술 개발 및 스마트화 현황 ===
slide8_content = [
    "[AI 기반 관제시스템]",
    "• 실시간 수질 예측 및 최적 운전",
    "• 이상 징후 조기 발견",
    "• SK에코플랜트 'AI Water', 삼성C&T '스마트워터'",
    "",
    "[에너지 자립화]",
    "• 슬러지 소화가스 발전",
    "• 태양광·수열에너지 활용",
    "• 에너지 소비 30% 절감 목표",
    "",
    "[디지털 트윈]",
    "• 가상 공간에서 시뮬레이션",
    "• 예방정비 및 효율 최적화",
    "• 대우건설, 현대건설 시범 적용 중"
]
add_content_slide(prs, "기술 개발 및 스마트 하수처리장 현황", slide8_content)

# === 슬라이드 9: 향후 시장 전망 및 과제 ===
slide9_content = [
    "[시장 전망]",
    "• 2025년 시장 규모: 1조 5,000억원 예상",
    "• 대형사 중심 재편 가속화 (Top 5 점유율 50%↑)",
    "• M&A를 통한 사업 portfolio 확대",
    "",
    "[주요 과제]",
    "• 고령화 처리시설 개보수 수요 증가",
    "• 미량오염물질 (약물, 내분비계) 처리 기술",
    "• 탄소중립 대응 에너지 효율화",
    "• 전문 인력 부족 및 처우 개선",
    "",
    "[ESG 경영]",
    "• 온실가스 감축, 재생에너지 확대",
    "• 지역사회 협력 프로그램 강화"
]
add_content_slide(prs, "향후 시장 전망 및 주요 과제", slide9_content)

# === 슬라이드 10: 결론 및 제언 ===
slide10_content = [
    "[종합 요약]",
    "• 전국 하수처리장의 88%가 민간 관리대행 운영",
    "• 상위 5개사가 시장의 45~50% 점유",
    "• 기술력·운영실적이 입찰 성패 결정",
    "",
    "[성공 전략 제언]",
    "1. AI·디지털 기술 선제적 도입",
    "2. 에너지 자립률 제고를 통한 비용경쟁력 확보",
    "3. 지역 사회와의 상생 협력 모델 구축",
    "4. 전문 인력 양성 및 체계적 교육",
    "5. ESG 경영 실천을 통한 기업 이미지 제고",
    "",
    "[마무리]",
    "• 지속가능한 하수처리 산업 생태계 조성 필요",
    "• 민관 협력을 통한 수질 환경 개선"
]
add_content_slide(prs, "결론 및 제언", slide10_content)

# 프레젠테이션 저장
output_path = "/workspace/하수처리장_관리대행_업체_현황.pptx"
prs.save(output_path)

print(f"프레젠테이션이 성공적으로 생성되었습니다: {output_path}")
print("총 10장의 슬라이드가 포함되었습니다.")

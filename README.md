# WURI Template Auto-Filler

WURI Foundation 혁신 프로그램 템플릿에 크롤링된 데이터를 자동으로 채워넣는 도구

## 요구사항

- Python 3.8+
- 필요 라이브러리: `python-docx`, `docxtpl`

## 설치

```bash
# 저장소 클론
git clone https://github.com/YOUR_USERNAME/wuri-template-filler.git
cd wuri-template-filler

# 의존성 설치
pip install -r requirements.txt
```

## 사용법

### 1. 파일 준비
- `Data.docx`: 크롤링된 데이터 (WURI 형식으로 정리된 상태)
- `template_original.docx`: WURI 빈 템플릿

### 2. 실행
```bash
python fill_wuri_template.py
```

### 3. 결과
- `WURI_Filled_YYYYMMDD_HHMMSS.docx` 파일 생성

## 파일 구조

```
wuri-template-filler/
├── README.md
├── requirements.txt
├── fill_wuri_template.py      # 메인 스크립트
├── Data.docx                  # 입력 데이터 (예시)
├── template_original.docx     # WURI 템플릿
└── template_with_placeholders.docx  # 자동 생성됨
```

## 작동 원리

1. **데이터 추출**: `Data.docx`에서 정규식으로 필드별 데이터 파싱
2. **템플릿 준비**: 원본 템플릿에 `{{변수명}}` 플레이스홀더 추가
3. **데이터 매핑**: `docxtpl` (Jinja2) 라이브러리로 데이터 채우기
4. **결과 저장**: 타임스탬프 포함된 새 파일 생성

## 지원 필드

| 섹션 | 필드 |
|------|------|
| Writer's Profile | university_name, writer_full_name, writer_title, writer_email |
| Program Profile | program_name, category, champion, school_college |
| Planning | long_term_goals, short_term_targets, rationale, initiators, ... |
| Doing | launch_date, program_content, progress, ... |
| Seeing | impacts_students, responses_industry, measurable_output, ... |
| Future | future_planning |

## 커스터마이징

### 새 필드 추가
`fill_wuri_template.py`의 `extract_data_from_docx()` 함수에서:
```python
data['new_field'] = extract_section(full_text, r'패턴(.+?)(?=종료패턴|$)')
```

### 플레이스홀더 수동 편집
`template_with_placeholders.docx`를 Word에서 열어 `{{변수명}}` 위치 조정 가능

## 라이선스

MIT License

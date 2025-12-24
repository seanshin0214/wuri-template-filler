# WURI Template Auto-Filler

WURI Foundation 혁신 프로그램 템플릿에 크롤링된 데이터를 자동으로 채워넣는 도구

---

## 설치 가이드 (최초 1회)

### 필요한 프로그램

| 프로그램 | 필수 여부 | 설치 방법 |
|---------|----------|----------|
| **Python 3.8+** | 필수 | https://www.python.org/downloads/ |
| **python-docx** | 필수 | `pip install python-docx` |
| **docxtpl** | 필수 | `pip install docxtpl` |
| Microsoft Word | 권장 | 결과 파일 열람용 (없으면 한글, 구글독스로 가능) |

### 설치 순서

1. **Python 설치**
   - https://www.python.org/downloads/ 에서 다운로드
   - 설치 시 **"Add Python to PATH"** 체크 필수!

2. **라이브러리 설치** (명령 프롬프트 또는 PowerShell에서)
   ```
   pip install python-docx docxtpl
   ```

3. **완료** - 이제 `WURI_실행.bat` 더블클릭 가능

---

## 사용 가이드

### 1단계: 폴더 구조 확인

```
WURI_Automation/
├── WURI_실행.bat          ← 이거 더블클릭!
├── Data.docx              ← 여기에 크롤링 데이터 넣기
├── template_original.docx ← WURI 빈 템플릿 (수정 X)
├── fill_wuri_template.py  ← 자동화 스크립트 (수정 X)
└── ...
```

---

### 2단계: Data.docx 준비

`Data.docx`를 열면 이런 형식입니다:

```
Writer's Profile
├── University name: MIT
├── Full name: (작성자 이름)
├── Email address: ...

Program Profile
├── Program name: Amogy (Startup)
├── Category: C7: University-Based...
├── School/College: ...

Summary of Program
├── Abstract: ...

Details of Program
├── 1. Planning
│   ├── Long-term Goals: ...
│   ├── Short-term Targets: ...
├── 2. Doing
│   ├── Launch date: ...
├── 3. Seeing
│   ├── Impacts on students: ...
└── 4. Future Planning
```

**새 프로그램 데이터로 바꾸려면:**
- `Data.docx` 열기
- 각 항목에 새로운 크롤링 데이터 붙여넣기
- 저장

---

### 3단계: 실행

`WURI_실행.bat` **더블클릭**

```
============================================================
  WURI Template Auto-Filler
============================================================

[1/3] Data.docx에서 데이터 추출 중...
[2/3] 플레이스홀더 템플릿 확인 중...
[3/3] 템플릿에 데이터 채우는 중...

============================================================
  완료! 결과 파일이 생성되었습니다.
============================================================
```

---

### 4단계: 결과 확인

폴더에 새 파일 생성됨:
```
WURI_Filled_20251224_180000.docx  ← 이게 완성본!
```

더블클릭해서 Word로 열고 확인

---

## 요약

1. `Data.docx`에 데이터 준비
2. `WURI_실행.bat` 더블클릭
3. `WURI_Filled_*.docx` 확인

---

## 지원 필드

| 섹션 | 필드 |
|------|------|
| Writer's Profile | university_name, writer_full_name, writer_title, writer_email |
| Program Profile | program_name, category, champion, school_college |
| Planning | long_term_goals, short_term_targets, rationale, initiators, ... |
| Doing | launch_date, program_content, progress, ... |
| Seeing | impacts_students, responses_industry, measurable_output, ... |
| Future | future_planning |

---

## 라이선스

MIT License

# tools

DB 마이그레이션 보조 도구 모음

---

## excel_to_comment_sql.py

테이블명세서 엑셀 파일을 읽어서 `ALTER TABLE ... MODIFY COLUMN ... COMMENT` SQL을 일괄 생성하는 스크립트입니다.

### 요구사항

```bash
pip install openpyxl pandas
```

### 사용법

```bash
# 1. 먼저 엑셀 구조 확인 (dry-run)
python3 excel_to_comment_sql.py --input 명세서.xlsx --dry-run

# 2. SQL 생성 후 파일로 저장
python3 excel_to_comment_sql.py --input 명세서.xlsx --output comments.sql

# 3. 특정 시트 지정
python3 excel_to_comment_sql.py --input 명세서.xlsx --sheet "컬럼정의" --output comments.sql

# 4. 헤더 행이 2행인 경우 (0-based)
python3 excel_to_comment_sql.py --input 명세서.xlsx --header-row 1 --output comments.sql
```

### 엑셀 컬럼 매핑 설정

스크립트 상단의 `COLUMN_MAP` 을 실제 명세서 헤더명에 맞게 수정합니다.

```python
COLUMN_MAP = {
    "table_phys":  "테이블명",       # 물리 테이블명 (필수)
    "table_logic": "테이블논리명",    # 테이블 한글명
    "col_phys":    "컬럼명",         # 물리 컬럼명 (필수)
    "col_logic":   "컬럼논리명",      # 컬럼 한글명
    "data_type":   "데이터타입",      # 컬럼 타입
    "length":      "길이",           # 길이/정밀도
    "nullable":    "NULL여부",       # Y/N or NULL/NOT NULL
    "pk":          "PK",            # PK 여부
    "default_val": "기본값",         # DEFAULT 값
    "comment":     "설명",          # COMMENT 전용 컬럼 (있으면 col_logic보다 우선)
}
```

### COMMENT 값 결정 우선순위

`comment` 컬럼 → `col_logic` 컬럼 → `table_logic` 컬럼 순으로 fallback 됩니다.  
명세서에 별도 "설명" 컬럼이 없으면 `COLUMN_MAP["comment"]` 에 `None` 을 지정하세요.

### 출력 예시

```sql
ALTER TABLE `ACQTX_CALC_TRACE`
    MODIFY COLUMN `TMPL_TEXT` TEXT NOT NULL COMMENT '기호식(스냅샷)',
    MODIFY COLUMN `SUBST_TEXT` TEXT NOT NULL COMMENT '대입식(스냅샷)·DISP 표시문 전문';

ALTER TABLE `ACQTX_HEADER`
    MODIFY COLUMN `MEMBER_ID` VARCHAR(20) NOT NULL COMMENT '회원ID',
    MODIFY COLUMN `TAX_YEAR` CHAR(4) NOT NULL COMMENT '과세연도',
    MODIFY COLUMN `ACQ_DATE` DATE NULL COMMENT '취득일자',
    MODIFY COLUMN `STD_AMT` DECIMAL(18,0) NOT NULL COMMENT '과세표준액';
```

### 주요 옵션

| 옵션 | 설명 |
|---|---|
| `--input` / `-i` | 엑셀 파일 경로 (필수) |
| `--output` / `-o` | 출력 SQL 파일 경로 (생략 시 stdout) |
| `--dry-run` / `-d` | SQL 생성 없이 엑셀 컬럼 구조만 출력 |
| `--sheet` / `-s` | 시트 이름 (기본: 첫 번째 시트) |
| `--header-row` | 헤더 행 번호 0-based (기본: 0) |

### 타입 포함 여부 설정

```python
# 스크립트 상단 설정값
INCLUDE_TYPE_IN_MODIFY = True   # MODIFY COLUMN에 타입 포함 (기본)
INCLUDE_TYPE_IN_MODIFY = False  # COMMENT만 변경 (가장 안전, 타입 정보 불확실할 때)
```

> **주의:** `INCLUDE_TYPE_IN_MODIFY = True` 일 때 타입이 잘못 생성되면 데이터 손실 위험이 있습니다.  
> 처음에는 `False` 로 실행하거나 개발 환경에서 먼저 검증하세요.

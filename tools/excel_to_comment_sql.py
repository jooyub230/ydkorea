"""
테이블명세서 엑셀 → ALTER TABLE COMMENT SQL 생성기

사용법:
    python excel_to_comment_sql.py --input 명세서.xlsx --output comments.sql
    python excel_to_comment_sql.py --input 명세서.xlsx --output comments.sql --dry-run

엑셀 컬럼 매핑 설정 (아래 COLUMN_MAP 을 실제 명세서에 맞게 수정):
    - 헤더가 있는 행 번호: HEADER_ROW
    - 각 항목이 있는 컬럼 헤더명: COLUMN_MAP
"""

import argparse
import sys
from pathlib import Path

import openpyxl
import pandas as pd

# ─────────────────────────────────────────────
# 설정값: 실제 엑셀 명세서에 맞게 수정하세요
# ─────────────────────────────────────────────

# 헤더가 있는 행 (0-based). 보통 0(1행) 또는 1(2행)
HEADER_ROW = 0

# 엑셀에서 읽을 시트 이름 (None 이면 첫 번째 시트)
SHEET_NAME = None

# 엑셀 컬럼 헤더명 매핑
# 키: 스크립트 내부 식별자 / 값: 실제 엑셀 헤더 문자열 (없으면 None)
COLUMN_MAP = {
    "table_phys":  "table_name",      # 물리 테이블명 (영문)           예: ytable1
    "table_logic": "table_comment",   # 테이블 한글명                   예: 납세자정보
    "col_phys":    "column_name",     # 물리 컬럼명                     예: field1
    "col_logic":   None,              # 컬럼 한글명 별도 컬럼 없음
    "data_type":   "column_type",     # 전체 타입 문자열                 예: varchar(6), double
    "length":      None,              # 길이는 column_type 에 포함되어 있음
    "nullable":    "is_nullable",     # NULL 허용 여부                  YES / NO
    "pk":          "column_key",      # 키 타입                        PRI / MUL
    "default_val": "column_default",  # 기본값
    "comment":     "column_comment",  # COMMENT 로 사용할 컬럼          예: 납세자번호
}

# COMMENT 로 사용할 컬럼 우선순위: "comment" → "col_logic" 순서로 fallback
COMMENT_PRIORITY = ["comment", "col_logic", "table_logic"]

# NULL 허용으로 간주할 값 목록 (대소문자 무시)
# is_nullable = YES → NULL 허용 / NO → NOT NULL
NULLABLE_VALUES = {"y", "yes", "null", "nullable", "true", "1", ""}

# 타입 재구성 여부
# True  → column_type + is_nullable 로 MODIFY COLUMN 에 타입 포함 (권장)
# False → COMMENT 변경만 생성 (타입 정보 생략)
INCLUDE_TYPE_IN_MODIFY = True

# ─────────────────────────────────────────────


def normalize(val) -> str:
    """셀 값을 문자열로 정규화"""
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def resolve_col(df: pd.DataFrame, key: str) -> str | None:
    """COLUMN_MAP 에서 실제 컬럼 헤더를 찾아 반환"""
    mapped = COLUMN_MAP.get(key)
    if mapped and mapped in df.columns:
        return mapped
    return None


def get_comment(row, cols: dict) -> str:
    """우선순위에 따라 COMMENT 값 추출"""
    for priority_key in COMMENT_PRIORITY:
        col = cols.get(priority_key)
        if col:
            val = normalize(row.get(col, ""))
            if val:
                return val
    return ""


def normalize_type(dtype: str) -> str:
    """MySQL 8.0 deprecated 타입 표기 정규화.

    - INT(11), BIGINT(20) 등 정수형 display width 제거  → INT, BIGINT
    - DOUBLE(10,2), FLOAT(8,2) 등 부동소수점 정밀도 제거 → DOUBLE, FLOAT
    - TINYINT(1) 은 boolean 관용 표기이므로 그대로 유지
    - DECIMAL(18,2), VARCHAR(50) 등 의미 있는 정밀도/길이는 유지
    """
    import re

    upper = dtype.upper().strip()

    # 정수형: TINYINT(1) 은 유지, 나머지 정수형 display width 제거
    int_types = r"^(SMALLINT|MEDIUMINT|INT|INTEGER|BIGINT)(\(\d+\))(\s+UNSIGNED)?$"
    m = re.match(int_types, upper)
    if m:
        unsigned = m.group(3) or ""
        return f"{m.group(1)}{unsigned.rstrip()}"

    # 부동소수점 정밀도 제거: DOUBLE(M,D), FLOAT(M,D)
    float_types = r"^(DOUBLE|FLOAT)(\(\d+,\d+\))(\s+UNSIGNED)?$"
    m = re.match(float_types, upper)
    if m:
        unsigned = m.group(3) or ""
        return f"{m.group(1)}{unsigned.rstrip()}"

    return dtype.upper()


def build_type_str(row, cols: dict) -> str:
    """데이터 타입 + 길이 문자열 조합"""
    dtype = normalize(row.get(cols.get("data_type", ""), ""))
    length = normalize(row.get(cols.get("length", ""), ""))

    if not dtype:
        return ""

    # 길이가 이미 타입에 포함된 경우 (VARCHAR(100) 형태)
    if "(" in dtype:
        return normalize_type(dtype)

    # 길이 정보가 별도 컬럼에 있는 경우
    if length and length not in ("0", "-"):
        return normalize_type(f"{dtype}({length})")

    return normalize_type(dtype)


def is_nullable(row, cols: dict) -> bool | None:
    """NULL 허용 여부 판단. 컬럼이 없으면 None 반환 (타입 문자열에 포함하지 않음)"""
    col = cols.get("nullable")
    if not col:
        return None
    val = normalize(row.get(col, "")).lower()
    return val in NULLABLE_VALUES


def escape_comment(text: str) -> str:
    """COMMENT 문자열 정제 및 이스케이프.
    - 셀 내 줄바꿈(Alt+Enter 등) → 공백으로 치환
    - 작은따옴표 이스케이프
    """
    text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    text = text.replace("'", "\\'")
    return text.strip()


# TEXT/BLOB 계열은 DEFAULT 값을 지정할 수 없음 (MySQL 제약)
_NO_DEFAULT_TYPES = {"tinytext", "text", "mediumtext", "longtext",
                     "tinyblob", "blob", "mediumblob", "longblob", "json"}


def build_default(row, cols: dict) -> str:
    """DEFAULT 절 문자열 반환. 값이 없거나 TEXT/BLOB 타입이면 빈 문자열."""
    col = cols.get("default_val")
    if not col:
        return ""
    val = normalize(row.get(col, ""))
    if not val or val.upper() == "NAN":
        return ""

    # TEXT/BLOB/JSON 계열은 DEFAULT 불가
    type_col = cols.get("data_type")
    if type_col:
        base_type = normalize(row.get(type_col, "")).lower().split("(")[0].strip()
        if base_type in _NO_DEFAULT_TYPES:
            return ""

    # NULL 키워드는 따옴표 없이
    if val.upper() == "NULL":
        return " DEFAULT NULL"
    # 숫자면 따옴표 없이
    try:
        float(val)
        return f" DEFAULT {val}"
    except ValueError:
        return f" DEFAULT '{escape_comment(val)}'"


def generate_sql(df: pd.DataFrame) -> list[str]:
    """데이터프레임 → ALTER TABLE SQL 목록 생성"""

    # 사용 가능한 컬럼 해석
    cols = {key: resolve_col(df, key) for key in COLUMN_MAP}

    table_col = cols.get("table_phys")
    col_col = cols.get("col_phys")

    if not table_col:
        print(f"[ERROR] 물리 테이블명 컬럼을 찾을 수 없습니다. COLUMN_MAP['table_phys'] 확인: '{COLUMN_MAP['table_phys']}'")
        print(f"        실제 엑셀 컬럼 목록: {list(df.columns)}")
        sys.exit(1)

    if not col_col:
        print(f"[ERROR] 물리 컬럼명 컬럼을 찾을 수 없습니다. COLUMN_MAP['col_phys'] 확인: '{COLUMN_MAP['col_phys']}'")
        sys.exit(1)

    # 테이블별로 그룹핑
    groups: dict[str, list[str]] = {}
    skipped = 0

    for _, row in df.iterrows():
        table = normalize(row.get(table_col, ""))
        column = normalize(row.get(col_col, ""))

        if not table or not column:
            skipped += 1
            continue

        comment = get_comment(row, cols)

        if INCLUDE_TYPE_IN_MODIFY:
            type_str = build_type_str(row, cols)
            nullable = is_nullable(row, cols)
            default_val = build_default(row, cols)

            # extra 컬럼에서 AUTO_INCREMENT 여부 확인
            extra_col = cols.get("pk")  # column_key 재사용 대신 extra 별도 처리
            extra_val = ""
            # COLUMN_MAP 에 "extra" 키가 있으면 사용, 없으면 "extra" 헤더 직접 탐색
            for candidate in ("extra",):
                if candidate in df.columns:
                    extra_val = normalize(row.get(candidate, "")).upper()
                    break
            auto_inc = " AUTO_INCREMENT" if "AUTO_INCREMENT" in extra_val else ""

            if type_str:
                null_part = ""
                if nullable is True:
                    null_part = " NULL"
                elif nullable is False:
                    null_part = " NOT NULL"
                modify_line = (
                    f"    MODIFY COLUMN `{column}` {type_str}{null_part}"
                    f"{default_val}{auto_inc} COMMENT '{escape_comment(comment)}'"
                )
            else:
                modify_line = f"    MODIFY COLUMN `{column}` COMMENT '{escape_comment(comment)}'"
        else:
            modify_line = f"    MODIFY COLUMN `{column}` COMMENT '{escape_comment(comment)}'"

        groups.setdefault(table, []).append(modify_line)

    if skipped:
        print(f"[INFO] 테이블명 또는 컬럼명이 비어있어 건너뛴 행: {skipped}개")

    # SQL 생성
    sql_blocks = []
    for table, lines in groups.items():
        # 스키마명이 포함된 경우(예: propert4x_db.ytable10) 각각 백틱으로 감쌈
        if "." in table:
            schema, tbl = table.split(".", 1)
            table_ref = f"`{schema}`.`{tbl}`"
        else:
            table_ref = f"`{table}`"
        block = f"ALTER TABLE {table_ref}\n" + ",\n".join(lines) + ";\n"
        sql_blocks.append(block)

    return sql_blocks


def read_excel(path: Path) -> pd.DataFrame:
    sheet = SHEET_NAME if SHEET_NAME is not None else 0
    result = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW, dtype=str)
    # sheet_name이 정수이거나 단일 시트명이면 DataFrame, 아니면 dict
    if isinstance(result, dict):
        df = next(iter(result.values()))
    else:
        df = result
    # 컬럼명 앞뒤 공백 제거
    df.columns = [str(c).strip() for c in df.columns]
    return df


def print_column_preview(df: pd.DataFrame):
    """엑셀 컬럼 목록과 첫 3행 미리보기 출력"""
    print("\n[ 엑셀 컬럼 목록 ]")
    for i, col in enumerate(df.columns):
        print(f"  {i:>3}: {col}")
    print("\n[ 데이터 미리보기 (상위 3행) ]")
    print(df.head(3).to_string())
    print()


def main():
    parser = argparse.ArgumentParser(
        description="테이블명세서 엑셀 → ALTER TABLE COMMENT SQL 생성기"
    )
    parser.add_argument("--input",   "-i", required=True,  help="엑셀 파일 경로")
    parser.add_argument("--output",  "-o", default=None,   help="출력 SQL 파일 경로 (생략 시 stdout)")
    parser.add_argument("--dry-run", "-d", action="store_true", help="SQL 생성 없이 엑셀 구조만 미리보기")
    parser.add_argument("--sheet",   "-s", default=None,   help="시트 이름 (기본: 첫 번째 시트)")
    parser.add_argument("--header-row", type=int, default=None, help="헤더 행 번호 0-based (기본: 설정값)")
    args = parser.parse_args()

    # 인자로 설정 오버라이드
    if args.sheet:
        global SHEET_NAME
        SHEET_NAME = args.sheet
    if args.header_row is not None:
        global HEADER_ROW
        HEADER_ROW = args.header_row

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[ERROR] 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    print(f"[INFO] 파일 읽는 중: {input_path}")
    df = read_excel(input_path)
    print(f"[INFO] 로드 완료: {len(df)}행, {len(df.columns)}컬럼")

    if args.dry_run:
        print_column_preview(df)
        # 현재 COLUMN_MAP 매핑 상태 확인
        _df_check = df

        print("[ COLUMN_MAP 매핑 상태 ]")
        all_ok = True
        for key, mapped in COLUMN_MAP.items():
            if mapped is None:
                print(f"  {key:15}: (사용 안 함)")
            elif mapped in _df_check.columns:
                print(f"  {key:15}: '{mapped}' ✓")
            else:
                print(f"  {key:15}: '{mapped}' ✗  ← 엑셀에 없는 컬럼명!")
                all_ok = False

        if all_ok:
            print("\n[DRY-RUN] 매핑 확인 완료. --dry-run 을 제거하고 다시 실행하면 SQL 이 생성됩니다.")
        else:
            print("\n[DRY-RUN] ✗ 표시된 항목의 COLUMN_MAP 값을 엑셀 컬럼명에 맞게 수정하세요.")
        return

    sql_blocks = generate_sql(df)
    total_tables = len(sql_blocks)
    total_cols = sum(b.count("MODIFY COLUMN") for b in sql_blocks)

    output_text = "\n".join(sql_blocks)

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(output_text, encoding="utf-8")
        print(f"[INFO] SQL 저장 완료: {output_path}")
    else:
        print(output_text)

    print(f"[INFO] 처리 완료: 테이블 {total_tables}개, 컬럼 {total_cols}개")


if __name__ == "__main__":
    main()

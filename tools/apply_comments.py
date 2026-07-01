"""
테이블명세서 엑셀 + DB 직접 연결 → COMMENT 정확 적용 스크립트

엑셀에서는 "어떤 컬럼에 어떤 COMMENT를 달 것인가"만 읽고,
실제 컬럼 정의(타입/길이/NULL/DEFAULT/EXTRA)는 DB의 INFORMATION_SCHEMA에서 직접 읽습니다.
→ 타입 불일치 오류 원천 차단

사용법:
    # SQL 파일만 생성 (DB 미접속)
    python3 apply_comments.py --input 명세서.xlsx --host localhost --user root --password secret --database mydb --output comments.sql

    # SQL 생성 + 즉시 DB 적용
    python3 apply_comments.py --input 명세서.xlsx --host localhost --user root --password secret --database mydb --apply

    # 적용 전 변경 내역 미리보기
    python3 apply_comments.py --input 명세서.xlsx --host localhost --user root --password secret --database mydb --dry-run
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────
# 설정값
# ─────────────────────────────────────────────

HEADER_ROW = 0
SHEET_NAME = None

# 엑셀 컬럼 헤더 매핑 (table_name, column_name, column_comment 는 필수)
COLUMN_MAP = {
    "table_phys": "table_name",      # 물리 테이블명  예: propert4x_db.ytable1
    "col_phys":   "column_name",     # 물리 컬럼명    예: field1
    "comment":    "column_comment",  # COMMENT 값     예: 납세자번호
}

# ─────────────────────────────────────────────


def normalize(val) -> str:
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def escape_comment(text: str) -> str:
    text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return text.replace("'", "\\'").strip()


def read_excel(path: Path) -> pd.DataFrame:
    sheet = SHEET_NAME if SHEET_NAME is not None else 0
    result = pd.read_excel(path, sheet_name=sheet, header=HEADER_ROW, dtype=str)
    if isinstance(result, dict):
        df = next(iter(result.values()))
    else:
        df = result
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_comment_map(df: pd.DataFrame) -> dict[tuple[str, str, str], str]:
    """엑셀에서 {(schema, table, column): comment} 딕셔너리 반환"""
    table_col   = COLUMN_MAP["table_phys"]
    col_col     = COLUMN_MAP["col_phys"]
    comment_col = COLUMN_MAP["comment"]

    for required in (table_col, col_col, comment_col):
        if required not in df.columns:
            print(f"[ERROR] 엑셀에서 '{required}' 컬럼을 찾을 수 없습니다.")
            print(f"        실제 컬럼 목록: {list(df.columns)}")
            sys.exit(1)

    mapping: dict[tuple[str, str, str], str] = {}
    for _, row in df.iterrows():
        table_full = normalize(row.get(table_col, ""))
        column     = normalize(row.get(col_col, ""))
        comment    = normalize(row.get(comment_col, ""))

        if not table_full or not column:
            continue

        # schema.table 분리
        if "." in table_full:
            schema, table = table_full.split(".", 1)
        else:
            schema, table = "", table_full

        mapping[(schema, table, column)] = comment

    return mapping


def fetch_column_definitions(conn, schema: str, tables: list[str]) -> dict[tuple[str, str], dict]:
    """INFORMATION_SCHEMA.COLUMNS에서 현재 컬럼 정의를 읽어옴"""
    import pymysql.cursors

    placeholders = ",".join(["%s"] * len(tables))
    query = f"""
        SELECT
            TABLE_SCHEMA,
            TABLE_NAME,
            COLUMN_NAME,
            COLUMN_TYPE,
            IS_NULLABLE,
            COLUMN_DEFAULT,
            EXTRA,
            COLUMN_COMMENT
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = %s
          AND TABLE_NAME IN ({placeholders})
        ORDER BY TABLE_NAME, ORDINAL_POSITION
    """
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute(query, [schema] + tables)
        rows = cur.fetchall()

    result = {}
    for r in rows:
        key = (r["TABLE_NAME"], r["COLUMN_NAME"])
        result[key] = r
    return result


def build_column_def(col_info: dict) -> str:
    """INFORMATION_SCHEMA 행 → MODIFY COLUMN 타입 정의 문자열"""
    parts = [col_info["COLUMN_TYPE"].upper()]

    if col_info["IS_NULLABLE"] == "NO":
        parts.append("NOT NULL")
    else:
        parts.append("NULL")

    default = col_info["COLUMN_DEFAULT"]
    col_type_base = col_info["COLUMN_TYPE"].lower().split("(")[0]
    no_default_types = {"tinytext","text","mediumtext","longtext",
                        "tinyblob","blob","mediumblob","longblob","json"}

    if default is not None and col_type_base not in no_default_types:
        if default.upper() == "NULL":
            parts.append("DEFAULT NULL")
        else:
            try:
                float(default)
                parts.append(f"DEFAULT {default}")
            except (ValueError, TypeError):
                parts.append(f"DEFAULT '{default}'")

    extra = (col_info["EXTRA"] or "").upper()
    if "AUTO_INCREMENT" in extra:
        parts.append("AUTO_INCREMENT")

    return " ".join(parts)


def generate_sql_blocks(
    comment_map: dict[tuple[str, str, str], str],
    col_defs: dict[tuple[str, str], dict],
    default_schema: str,
) -> tuple[list[str], list[str]]:
    """ALTER TABLE SQL 블록 목록과 경고 메시지 목록 반환"""

    # 테이블별로 그룹핑
    groups: dict[tuple[str, str], list[str]] = {}
    warnings = []

    for (schema, table, column), comment in comment_map.items():
        s = schema or default_schema
        col_info = col_defs.get((table, column))

        if col_info is None:
            warnings.append(f"[WARN] DB에서 찾을 수 없음: {s}.{table}.{column} (건너뜀)")
            continue

        col_def = build_column_def(col_info)
        line = f"    MODIFY COLUMN `{column}` {col_def} COMMENT '{escape_comment(comment)}'"
        groups.setdefault((s, table), []).append(line)

    sql_blocks = []
    for (s, table), lines in groups.items():
        block = f"ALTER TABLE `{s}`.`{table}`\n" + ",\n".join(lines) + ";\n"
        sql_blocks.append(block)

    return sql_blocks, warnings


def main():
    parser = argparse.ArgumentParser(
        description="DB INFORMATION_SCHEMA 기반 COMMENT 적용 스크립트"
    )
    parser.add_argument("--input",    "-i", required=True,        help="엑셀 파일 경로")
    parser.add_argument("--host",           default="localhost",   help="DB 호스트")
    parser.add_argument("--port",     type=int, default=3306,     help="DB 포트")
    parser.add_argument("--user",     "-u", required=True,        help="DB 사용자")
    parser.add_argument("--password", "-p", required=True,        help="DB 비밀번호")
    parser.add_argument("--database", "-d", required=True,        help="기본 스키마명")
    parser.add_argument("--output",   "-o", default=None,         help="SQL 출력 파일 경로 (생략 시 stdout)")
    parser.add_argument("--apply",          action="store_true",  help="SQL 즉시 DB 적용")
    parser.add_argument("--dry-run",        action="store_true",  help="변경 내역 미리보기만 출력")
    parser.add_argument("--sheet",    "-s", default=None,         help="시트 이름")
    args = parser.parse_args()

    if args.sheet:
        global SHEET_NAME
        SHEET_NAME = args.sheet

    # 1. 엑셀 읽기
    print(f"[INFO] 엑셀 읽는 중: {args.input}")
    df = read_excel(Path(args.input))
    print(f"[INFO] 로드 완료: {len(df)}행")

    comment_map = build_comment_map(df)
    print(f"[INFO] 컬럼 매핑 수: {len(comment_map)}개")

    # 2. DB 연결 및 INFORMATION_SCHEMA 조회
    import pymysql
    print(f"[INFO] DB 연결 중: {args.user}@{args.host}:{args.port}/{args.database}")
    try:
        conn = pymysql.connect(
            host=args.host, port=args.port,
            user=args.user, password=args.password,
            database="information_schema",
            charset="utf8mb4",
        )
    except Exception as e:
        print(f"[ERROR] DB 연결 실패: {e}")
        sys.exit(1)

    # 대상 테이블 목록 추출
    target_tables: dict[str, set[str]] = {}
    for schema, table, _ in comment_map:
        s = schema or args.database
        target_tables.setdefault(s, set()).add(table)

    col_defs: dict[tuple[str, str], dict] = {}
    for schema, tables in target_tables.items():
        print(f"[INFO] INFORMATION_SCHEMA 조회: {schema} ({len(tables)}개 테이블)")
        col_defs.update(fetch_column_definitions(conn, schema, list(tables)))

    print(f"[INFO] DB 컬럼 정의 로드: {len(col_defs)}개")

    # 3. SQL 생성
    sql_blocks, warnings = generate_sql_blocks(comment_map, col_defs, args.database)

    for w in warnings:
        print(w)

    total_cols = sum(b.count("MODIFY COLUMN") for b in sql_blocks)
    output_text = "\n".join(sql_blocks)

    if args.dry_run:
        print(f"\n[DRY-RUN] 적용 예정: 테이블 {len(sql_blocks)}개, 컬럼 {total_cols}개")
        print("\n" + output_text[:3000] + (" ..." if len(output_text) > 3000 else ""))
        conn.close()
        return

    # 4. SQL 파일 저장
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(output_text, encoding="utf-8")
        print(f"[INFO] SQL 저장 완료: {out}  (테이블 {len(sql_blocks)}개, 컬럼 {total_cols}개)")

    # 5. DB 직접 적용
    if args.apply:
        print(f"[INFO] DB 적용 시작...")
        ok = fail = 0
        with conn.cursor() as cur:
            for block in sql_blocks:
                try:
                    cur.execute(block)
                    conn.commit()
                    ok += 1
                except Exception as e:
                    conn.rollback()
                    print(f"[ERROR] 실패:\n{block.splitlines()[0]}\n  → {e}")
                    fail += 1
        print(f"[INFO] 완료: 성공 {ok}개, 실패 {fail}개")

    if not args.output and not args.apply:
        print(output_text)
        print(f"[INFO] 처리 완료: 테이블 {len(sql_blocks)}개, 컬럼 {total_cols}개")

    conn.close()


if __name__ == "__main__":
    main()

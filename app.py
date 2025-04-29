# app.py ───────────────────────────────────────────────────────────────
import io
import pathlib
import re

import pandas as pd
import streamlit as st
import xlsxwriter
import openpyxl         # ← requirements.txt 에 이미 명시돼 있어야 합니다

# ──────────────────────────────────────────────────────────────────────
# ① “S2 콘텐츠 전체”(file 3)는 리포지터리 **data/** 폴더에 고정된 파일을 사용
DATA_DIR   = pathlib.Path(__file__).parent / "data"
FILE3_PATH = DATA_DIR / "file3_default.xlsx"      # ← data/file3_default.xlsx
# ──────────────────────────────────────────────────────────────────────

# ── 자동‧다국어 제목 컬럼 후보 -------------------------------------------------
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = ["컨텐츠", "타이틀", "작품명", "도서명", "작품 제목",
                  "상품명", "이용상품명", "상품 제목", "ProductName", "Title", "제목"]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ── 공통 유틸 ------------------------------------------------------------------
def pick(candidates: list[str], df: pd.DataFrame) -> str:
    """가장 먼저 매칭되는 컬럼명을 돌려준다."""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {candidates}")

def clean_title(text: str) -> str:
    t = str(text)

    # “ 제숫자권/화 ” 패턴 제거
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)

    # 특수 치환
    for k, v in {
        "Un-holyNight": "UnholyNight",
        "?": "", "~": "", ",": "", "-": "", "_": ""
    }.items():
        t = t.replace(k, v)

    # 괄호·대괄호 내용 제거
    t = re.sub(r"\([^)]*\)", "", t)
    t = re.sub(r"\[[^\]]*\]", "", t)

    # 숫자+권/화/부/회
    t = re.sub(r"\d+[권화부회]", "", t)

    # 불필요 키워드
    for kw in [
        "개정판 l", "개정판", "외전", "무삭제본", "무삭제판", "합본",
        "단행본", "시즌", "세트", "연재", "특별", "최종화", "완결",
        "2부", "무삭제", "완전판", "세개정판", "19세개정판"
    ]:
        t = t.replace(kw, "")

    # 기타 노이즈
    t = re.sub(r"\d+", "", t).rstrip(".")
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}()]", "", t)
    t = re.sub(r"특별$", "", t)
    return t.replace(" ", "").strip()

# ── UI ------------------------------------------------------------------------
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

f1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
f2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")

st.write("③ **S2 콘텐츠 전체 (file3)** 는 리포지터리의 "
         "`data/file3_default.xlsx` 를 자동으로 사용합니다.")
save_name = st.text_input("💾 저장 파일명(확장자 생략 가능)", "mapping_result")
if not save_name.lower().endswith(".xlsx"):
    save_name += ".xlsx"

# ── 실행 ----------------------------------------------------------------------
if st.button("🟢 매핑 실행"):

    # 1) 입력‧파일 존재 확인 ----------------------------------------------------
    if not (f1 and f2):
        st.error("file1, file2 두 개의 엑셀을 먼저 업로드해 주세요.")
        st.stop()
    if not FILE3_PATH.exists():
        st.error(f"⚠️ 3번 파일이 {FILE3_PATH} 에 없습니다. "
                 "리포지터리에 data 폴더와 파일을 추가한 뒤 다시 실행해 주세요.")
        st.stop()

    # 2) 엑셀 → DataFrame ------------------------------------------------------
    df1 = pd.read_excel(f1)
    df2 = pd.concat(pd.read_excel(f2, sheet_name=None).values(),
                    ignore_index=True)
    df3 = pd.read_excel(FILE3_PATH)

    # 3) 컬럼 탐색 --------------------------------------------------------------
    c1  = pick(FILE1_COL_CAND, df1)
    c2  = pick(FILE2_COL_CAND, df2)
    c3  = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 4) 제목 정제 --------------------------------------------------------------
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 5) 1차 매핑 (file1 → file2) ----------------------------------------------
    map1 = (df1.drop_duplicates("정제_콘텐츠명")
              .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"])
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 6) 2차 매핑 (file3 → file2) ----------------------------------------------
    map3 = (df3.drop_duplicates("정제_콘텐츠3명")
              .set_index("정제_콘텐츠3명")[id3])
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 7) 매핑콘텐츠명 / 콘텐츠ID -----------------------------------------------
    mask_pair   = df2["정제_상품명"] == df2["매핑결과"]
    base_pairs  = (df2.loc[mask_pair, ["정제_상품명", "최종_매핑결과"]]
                      .query("`정제_상품명`.str.strip() != ''", engine="python")
                      .drop_duplicates()
                      .rename(columns={"정제_상품명": "매핑콘텐츠명",
                                       "최종_매핑결과": "콘텐츠ID"}))
    base_pairs["매핑콘텐츠명"] = base_pairs["매핑콘텐츠명"].apply(clean_title)

    same_mask     = base_pairs["매핑콘텐츠명"] == base_pairs["콘텐츠ID"]
    pairs_unique  = base_pairs.loc[~same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)
    pairs_same    = base_pairs.loc[ same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)

    pad_u = len(df2) - len(pairs_unique)
    df2["매핑콘텐츠명"] = list(pairs_unique["매핑콘텐츠명"]) + [""] * pad_u
    df2["콘텐츠ID"]     = list(pairs_unique["콘텐츠ID"])     + [""] * pad_u

    pad_s = len(df2) - len(pairs_same)
    df2["동일_매핑콘텐츠명"] = list(pairs_same["매핑콘텐츠명"]) + [""] * pad_s
    df2["동일_콘텐츠ID"]     = list(pairs_same["콘텐츠ID"])     + [""] * pad_s

    # 8) file1 정보 붙이기 ------------------------------------------------------
    info = (df1[[c1, "정제_콘텐츠명", "판매채널콘텐츠ID"]]
              .rename(columns={c1: "file1_콘텐츠명",
                               "정제_콘텐츠명": "file1_정제_콘텐츠명",
                               "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID"}))
    result = pd.concat([df2, info], axis=1)

    # 9) 열 순서 재배치 ---------------------------------------------------------
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    cols  = list(result.columns)
    idx   = cols.index("콘텐츠ID") + 1
    for col in ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]:
        cols.remove(col)
    cols[idx:idx] = ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]
    result = result[front + [c for c in cols if c not in front]]

    # 10) 엑셀 저장 & 헤더 색상 -------------------------------------------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

        wb = writer.book
        ws = writer.sheets["매핑결과"]

        fmt_y = wb.add_format({"bg_color": "#FFFFCC", "bold": True, "border": 1})
        fmt_g = wb.add_format({"bg_color": "#99FFCC", "bold": True, "border": 1})

        for col_idx, col_name in enumerate(result.columns):
            if col_name in {"매핑콘텐츠명", "콘텐츠ID"}:
                ws.write(0, col_idx, col_name, fmt_y)   # 노랑
            elif col_name in {"동일_매핑콘텐츠명", "동일_콘텐츠ID"}:
                ws.write(0, col_idx, col_name, fmt_g)   # 연두

    # 11) 다운로드 --------------------------------------------------------------
    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        buf.getvalue(),
        file_name=save_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

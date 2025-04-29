# app.py ───────────────────────────────────────────────────────────────
import streamlit as st, pandas as pd, re, io
import openpyxl, xlsxwriter        # requirements.txt 에 이미 명시

# ── 0. 후보 컬럼 ──────────────────────────────────────────────────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = [
    "컨텐츠","타이틀","작품명","도서명","작품 제목",
    "상품명","이용상품명","상품 제목","ProductName","Title","제목"
]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID","콘텐츠ID","ID","ContentID"]

# ── 1. 유틸 함수 ─────────────────────────────────────────────────────
def pick(candidates: list[str], df: pd.DataFrame) -> str:
    """DataFrame 에서 첫 번째로 매칭되는 컬럼명을 반환."""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {candidates}")

def clean_title(text: str) -> str:
    t = str(text)

    # ① “ 제숫자권/화 ” 제거
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)

    # ② 특수 치환
    for k, v in {
        "Un-holyNight": "UnholyNight",
        "?":  "", "~": "", ",": "", "-": "", "_": ""
    }.items():
        t = t.replace(k, v)

    # ③ 괄호/대괄호 제거
    t = re.sub(r"\([^)]*\)", "", t)
    t = re.sub(r"\[[^\]]*\]", "", t)

    # ④ 숫자+권/화/부/회
    t = re.sub(r"\d+[권화부회]", "", t)

    # ⑤ 키워드 제거
    for kw in [
        "개정판 l","개정판","외전","무삭제본","무삭제판","합본",
        "단행본","시즌","세트","연재","특별","최종화","완결",
        "2부","무삭제","완전판","세개정판","19세개정판"
    ]:
        t = t.replace(kw, "")

    # ⑥ 기타
    t = re.sub(r"\d+", "", t).rstrip(".")
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}()]", "", t)
    t = re.sub(r"특별$", "", t)
    t = t.replace("[", "").replace("]", "")

    return t.replace(" ", "").strip()

# ── 2. Streamlit UI ──────────────────────────────────────────────────
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

f1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
f2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
f3 = st.file_uploader("③ S2 콘텐츠 전체 (file3)", type="xlsx")

# ── 3. 매핑 실행 버튼 ────────────────────────────────────────────────
if st.button("🟢 매핑 실행"):

    # 3-1) 입력 체크 --------------------------------------------------
    if not (f1 and f2 and f3):
        st.error("3개의 엑셀 파일을 모두 업로드해 주세요.")
        st.stop()

    # 3-2) Excel → DataFrame ----------------------------------------
    df1 = pd.read_excel(f1)
    df2 = pd.concat(pd.read_excel(f2, sheet_name=None).values(), ignore_index=True)
    df3 = pd.read_excel(f3)

    # 3-3) 제목/ID 컬럼 자동 선택 ------------------------------------
    c1  = pick(FILE1_COL_CAND, df1)
    c2  = pick(FILE2_COL_CAND, df2)
    c3  = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 3-4) 제목 정제 --------------------------------------------------
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 3-5) 1차 매핑 (file1 → file2) ---------------------------------
    map1 = (
        df1.drop_duplicates("정제_콘텐츠명")
           .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    )
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 3-6) 2차 매핑 (file3 → file2) ---------------------------------
    map3 = (
        df3.drop_duplicates("정제_콘텐츠3명")
           .set_index("정제_콘텐츠3명")[id3]
    )
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 3-7) 매핑콘텐츠명 / 콘텐츠ID -----------------------------------
    mask_pair = df2["정제_상품명"] == df2["매핑결과"]

    base_pairs = (
        df2.loc[mask_pair, ["정제_상품명", "최종_매핑결과"]]
           .query("`정제_상품명`.str.strip() != ''", engine="python")
           .drop_duplicates()
           .rename(columns={"정제_상품명": "매핑콘텐츠명",
                            "최종_매핑결과": "콘텐츠ID"})
    )
    base_pairs["매핑콘텐츠명"] = base_pairs["매핑콘텐츠명"].apply(clean_title)

    same_mask      = base_pairs["매핑콘텐츠명"] == base_pairs["콘텐츠ID"]
    pairs_unique   = base_pairs.loc[~same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)
    pairs_same     = base_pairs.loc[same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)

    pad_u = len(df2) - len(pairs_unique)
    df2["매핑콘텐츠명"] = list(pairs_unique["매핑콘텐츠명"]) + [""] * pad_u
    df2["콘텐츠ID"]     = list(pairs_unique["콘텐츠ID"])     + [""] * pad_u

    pad_s = len(df2) - len(pairs_same)
    df2["동일_매핑콘텐츠명"] = list(pairs_same["매핑콘텐츠명"]) + [""] * pad_s
    df2["동일_콘텐츠ID"]     = list(pairs_same["콘텐츠ID"])     + [""] * pad_s

    # 3-8) 최종 미매핑 & 정렬 ---------------------------------------
    final_unmatch = (
        df2.loc[mask_pair, "정제_상품명"]
           .drop_duplicates()
           .pipe(lambda s: s[~s.isin(map3.index)])          # 2차 매핑 실패
           .pipe(lambda s: s[~s.isin(base_pairs["매핑콘텐츠명"])])  # 이미 사용된 제목 제외
    )

    df2["최종_정렬된_매핑되지않은_상품명"] = (
        sorted(final_unmatch) + [""] * (len(df2) - len(final_unmatch))
    )
    df2["최종_매핑되지않은_상품명"] = df2["정제_상품명"].where(
        df2["정제_상품명"].isin(final_unmatch), ""
    )

    # 3-9) file1 정보 결합 ------------------------------------------
    info = (
        df1[[c1, "정제_콘텐츠명", "판매채널콘텐츠ID"]]
           .rename(columns={
               c1: "file1_콘텐츠명",
               "정제_콘텐츠명": "file1_정제_콘텐츠명",
               "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID"
           })
    )
    result = pd.concat([df2, info], axis=1)

    # 3-10) 열 순서 재배치 ------------------------------------------
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    cols  = list(result.columns)
    idx   = cols.index("콘텐츠ID") + 1
    for col in ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]:
        cols.remove(col)
    cols[idx:idx] = ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]
    ordered = front + [c for c in cols if c not in front]
    result  = result[ordered]

    # 3-11) 필요 없는 열 삭제 ---------------------------------------
    result.drop(
        columns=["동일_콘텐츠ID", "최종_정렬된_매핑되지않은_상품명", "최종_매핑되지않은_상품명"],
        errors="ignore",
        inplace=True,
    )

    # 3-12) 엑셀 → 메모리 저장 + 헤더 색상 --------------------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # (1) 데이터 기록
        result.to_excel(writer, sheet_name="매핑결과", index=False)

        # (2) 워크북 / 워크시트
        wb = writer.book
        ws = writer.sheets["매핑결과"]

        # (3) 서식
        fmt_yellow = wb.add_format({"bg_color": "#FFFFCC", "bold": True, "border": 1})
        fmt_green  = wb.add_format({"bg_color": "#99FFCC", "bold": True, "border": 1})

        yellow_cols = {"매핑콘텐츠명", "콘텐츠ID"}
        green_cols  = {"동일_매핑콘텐츠명"}

        for col_idx, col_name in enumerate(result.columns):
            if col_name in yellow_cols:
                ws.write(0, col_idx, col_name, fmt_yellow)
            elif col_name in green_cols:
                ws.write(0, col_idx, col_name, fmt_green)

    # 3-13) 사용자에게 파일명 입력받기 -------------------------------
    st.success("✅ 매핑 완료!  저장할 파일명을 입력한 뒤 다운로드하세요.")

    file_label = st.text_input("💾 저장할 파일명 (.xlsx 제외)", value="mapping_result")
    save_name  = (file_label or "mapping_result").strip() + ".xlsx"

    # 3-14) 다운로드 버튼
    st.download_button(
        "📥 엑셀 다운로드",
        buf.getvalue(),
        file_name=save_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_btn",
    )

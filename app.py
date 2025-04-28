# app.py ───────────────────────────────────────────────────────────────
import streamlit as st, pandas as pd, re, io
import openpyxl, xlsxwriter       # requirements.txt 에 이미 명시

# ── 후보 컬럼 ──────────────────────────────────────────────────────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = ["컨텐츠", "타이틀", "작품명", "도서명", "작품 제목",
                  "상품명", "이용상품명", "상품 제목", "ProductName", "Title", "제목"]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ── 공통 유틸 ──────────────────────────────────────────────────────────
def pick(candidates: list[str], df: pd.DataFrame) -> str:
    """DataFrame 에서 가장 먼저 매칭되는 컬럼명을 리턴"""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {candidates}")

def clean_title(text: str) -> str:
    t = str(text)

    # “ 제숫자권/화 ” 제거
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)

    # 특수 치환
    for k, v in {
        "Un-holyNight": "UnholyNight",
        "?":  "", "~": "", ",": "", "-": "", "_": ""
    }.items():
        t = t.replace(k, v)

    # 괄호·대괄호 안 내용 제거
    t = re.sub(r"\([^)]*\)", "", t)
    t = re.sub(r"\[[^\]]*\]", "", t)

    # 숫자+권/화/부/회
    t = re.sub(r"\d+[권화부회]", "", t)

    # 키워드
    for kw in [
        "개정판 l", "개정판", "외전", "무삭제본", "무삭제판", "합본",
        "단행본", "시즌", "세트", "연재", "특별", "최종화", "완결",
        "2부", "무삭제", "완전판", "세개정판", "19세개정판"
    ]:
        t = t.replace(kw, "")

    # 기타
    t = re.sub(r"\d+", "", t).rstrip(".")
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}$begin:math:display$$end:math:display$()]", "", t)
    t = re.sub(r"특별$", "", t)
    t = t.replace("[", "").replace("]", "")
    return t.replace(" ", "").strip()

# ── Streamlit UI ──────────────────────────────────────────────────────
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

f1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
f2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
f3 = st.file_uploader("③ S2 콘텐츠 전체 (file3)", type="xlsx")

# ── 실행 버튼 ---------------------------------------------------------
if st.button("🟢 매핑 실행"):

    # 1) 입력 확인
    if not (f1 and f2 and f3):
        st.error("3개의 엑셀 파일을 모두 업로드해 주세요.")
        st.stop()

    # 2) Excel → DataFrame
    df1 = pd.read_excel(f1)
    df2 = pd.concat(pd.read_excel(f2, sheet_name=None).values(), ignore_index=True)
    df3 = pd.read_excel(f3)

    # 3) 제목/ID 컬럼 자동 선택
    c1  = pick(FILE1_COL_CAND, df1)
    c2  = pick(FILE2_COL_CAND, df2)
    c3  = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 4) 제목 정제
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 5) 1차 매핑 (file1 → file2)
    map1 = (df1.drop_duplicates("정제_콘텐츠명")
              .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"])
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 6) 2차 매핑 (file3 → file2)
    map3 = (df3.drop_duplicates("정제_콘텐츠3명")
              .set_index("정제_콘텐츠3명")[id3])
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 7) 매핑콘텐츠명 / 콘텐츠ID (둘이 같은 경우 ↔ 다른 경우 분리)
    mask_pair = df2["정제_상품명"] == df2["매핑결과"]

    base_pairs = (
        df2.loc[mask_pair, ["정제_상품명", "최종_매핑결과"]]
           .query("`정제_상품명`.str.strip() != ''", engine="python")
           .drop_duplicates()
           .rename(columns={"정제_상품명": "매핑콘텐츠명",
                            "최종_매핑결과": "콘텐츠ID"})
    )
    # 🔄 한 번 더 정제해 혹시 남은 ‘[’ 등 제거
    base_pairs["매핑콘텐츠명"] = base_pairs["매핑콘텐츠명"].apply(clean_title)

    same_mask      = base_pairs["매핑콘텐츠명"] == base_pairs["콘텐츠ID"]
    pairs_unique   = base_pairs.loc[~same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)
    pairs_same     = base_pairs.loc[same_mask].sort_values("매핑콘텐츠명").reset_index(drop=True)

    # 7-B) df2 에 채워 넣기
    pad_unique = len(df2) - len(pairs_unique)
    df2["매핑콘텐츠명"] = list(pairs_unique["매핑콘텐츠명"]) + [""] * pad_unique
    df2["콘텐츠ID"]     = list(pairs_unique["콘텐츠ID"])     + [""] * pad_unique

    pad_same = len(df2) - len(pairs_same)
    df2["동일_매핑콘텐츠명"] = list(pairs_same["매핑콘텐츠명"]) + [""] * pad_same
    df2["동일_콘텐츠ID"]     = list(pairs_same["콘텐츠ID"])     + [""] * pad_same

    # 8) 최종 미매핑 & 정렬
    final_unmatch = (
        df2.loc[mask_pair, "정제_상품명"]
           .drop_duplicates()
           .pipe(lambda s: s[~s.isin(map3.index)])      # 2차 매핑도 실패한 것
           .pipe(lambda s: s[~s.isin(base_pairs["매핑콘텐츠명"])])  # 이미 사용된 제목 제외
    )

    df2["최종_정렬된_매핑되지않은_상품명"] = (
        sorted(final_unmatch) + [""] * (len(df2) - len(final_unmatch))
    )
    df2["최종_매핑되지않은_상품명"] = df2["정제_상품명"].where(
        df2["정제_상품명"].isin(final_unmatch), ""
    )

    # 9) file1 정보 붙이기
    info = (df1[[c1, "정제_콘텐츠명", "판매채널콘텐츠ID"]]
              .rename(columns={c1: "file1_콘텐츠명",
                               "정제_콘텐츠명": "file1_정제_콘텐츠명",
                               "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID"}))
    result = pd.concat([df2, info], axis=1)

    # 10) 열 순서 재배치
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    cols  = list(result.columns)
    # ‘콘텐츠ID’ 바로 뒤에 동일_* 두 열 삽입
    idx = cols.index("콘텐츠ID") + 1
    for col in ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]:
        cols.remove(col)
    cols[idx:idx] = ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]
    ordered = front + [c for c in cols if c not in front]
    result  = result[ordered]
    )
  
    result.drop(
    columns=["동일_콘텐츠ID", "최종_정렬된_매핑되지않은_상품명", "최종_매핑되지않은_상품명"],
    errors="ignore",
    inplace=True,

    # 11) 결과 저장 & 다운로드
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        buffer.getvalue(),
        file_name="mapping_result.xlsx",                 # ← 필요하면 자유롭게 바꿔도 됩니다
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

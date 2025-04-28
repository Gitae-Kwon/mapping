# app.py
import streamlit as st
import pandas as pd
import re
import io
import openpyxl
import xlsxwriter

# ────────── 유틸 함수들 ──────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = ["컨텐츠", "타이틀", "작품명", "도서명", "작품 제목",
                  "상품명", "이용상품명", "상품 제목", "ProductName", "Title", "제목"]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

def pick(col_list, df):
    for c in col_list:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {col_list}")

def clean_title(text: str) -> str:
    t = str(text)
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)
    for k, v in {"Un-holyNight": "UnholyNight", "?" : "", "~": "", ",": "", "-": "", "_": ""}.items():
        t = t.replace(k, v)
    t = re.sub(r"\([^)]*\)", "", t)
    t = re.sub(r"\[[^\]]*\]", "", t)
    t = re.sub(r"\d+[권화부회]", "", t)
    for kw in ["개정판 l","개정판","외전","무삭제본","무삭제판","합본",
               "단행본","시즌","세트","연재","특별","최종화","완결",
               "2부","무삭제","완전판","세개정판","19세개정판"]:
        t = t.replace(kw, "")
    t = re.sub(r"\d+", "", t).rstrip('.')
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}$begin:math:display$$end:math:display$()]","",t)
    t = re.sub(r"특별$", "", t)
    return t.replace(" ", "").strip()

# ────────── Streamlit UI ──────────
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

file1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
file2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
file3 = st.file_uploader("③ S2 콘텐츠 전체 (file3)", type="xlsx")

if st.button("🟢 매핑 실행"):

    if not (file1 and file2 and file3):
        st.error("3개의 엑셀 파일을 모두 업로드해 주세요.")
        st.stop()

    # Excel → DataFrame
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2, sheet_name=None)     # file2는 시트 여러 개 가능
    df2 = pd.concat(df2.values(), ignore_index=True)

    df3 = pd.read_excel(file3)

    # 컬럼 선택
    c1 = pick(FILE1_COL_CAND, df1)
    c2 = pick(FILE2_COL_CAND, df2)
    c3 = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 정제
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 1차 매핑
    map1 = df1.drop_duplicates("정제_콘텐츠명").set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 2차 매핑
    map3 = df3.drop_duplicates("정제_콘텐츠3명").set_index("정제_콘텐츠3명")[id3]
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 미매핑 정렬 컬럼
    no1 = df2.loc[df2["정제_상품명"] == df2["매핑결과"], "정제_상품명"]
    final_unmatch = no1[~no1.isin(map3.index)].drop_duplicates()
    df2["최종_정렬된_매핑되지않은_상품명"] = (
        sorted(final_unmatch) + [""]*(len(df2)-len(final_unmatch))
    )

    # file1 정보 붙이기
    info = df1[[c1,"정제_콘텐츠명","판매채널콘텐츠ID"]].rename(columns={
        c1:"file1_콘텐츠명","정제_콘텐츠명":"file1_정제_콘텐츠명",
        "판매채널콘텐츠ID":"file1_판매채널콘텐츠ID"})
    result = pd.concat([df2, info], axis=1)

    # 엑셀로 메모리에 저장
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)
    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button("📥 결과 엑셀 다운로드", out.getvalue(),
                       file_name="mapping_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

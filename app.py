# app.py ───────────────────────────────────────────────────────────────
import pathlib, io, re
import streamlit as st
import pandas as pd
import openpyxl, xlsxwriter   # requirements.txt 에 이미 명시됨

# ── (고정) ③번 파일 경로 ------------------------------------------------
DATA_DIR   = pathlib.Path(__file__).parent / "data"
FILE3_PATH = DATA_DIR / "all_contents.xlsx"      # data/file3_default.xlsx

# ── 컬럼 후보 -----------------------------------------------------------
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = [
    "컨텐츠","타이틀","작품명","도서명","작품 제목",
    "상품명","이용상품명","상품 제목","ProductName","Title","제목"
]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ── 유틸 ----------------------------------------------------------------
def pick(candidates: list[str], df: pd.DataFrame) -> str:
    """DataFrame 안에서 후보 이름과 일치하는 첫 번째 컬럼명을 리턴"""
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {candidates}")

def clean_title(text: str) -> str:
    """제목 정제 규칙 한 곳에 모아두기"""
    t = str(text)
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)                       # ① ‘제 n권/화’
    for k, v in {"Un-holyNight": "UnholyNight", "?" : "", "~": "", ",": "",
                 "-" : "", "_" : ""}.items():
        t = t.replace(k, v)                                       # ② 특수 치환
    t = re.sub(r"\([^)]*\)|\[[^\]]*\]", "", t)                    # ③ 괄호·대괄호
    t = re.sub(r"\d+[권화부회]", "", t)                           # ④ 123권/화/부/회
    for kw in ["개정판 l","개정판","외전","무삭제본","무삭제판","합본",
               "단행본","시즌","세트","연재","특별","최종화","완결",
               "2부","무삭제","완전판","세개정판","19세개정판"]:
        t = t.replace(kw, "")
    t = re.sub(r"\d+", "", t).rstrip(".")                         # ⑤ 끝점·숫자
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}()]", "", t)
    t = re.sub(r"특별$", "", t)                                   # ⑥ ‘…특별’ 꼬리
    return t.replace(" ", "").strip()

# ── Streamlit UI --------------------------------------------------------
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

f1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
f2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
st.markdown("③ **S2 콘텐츠 전체 리스트는 IPS(4월28일기준)데이터를 자동으로 사용합니다.")

save_name = st.text_input("💾 저장 파일명(확장자 제외)", value="mapping_result")
save_name = (save_name if save_name.lower().endswith(".xlsx") else save_name + ".xlsx")

# ── 실행 ----------------------------------------------------------------
if st.button("🟢 매핑 실행"):

    # 1) 입력 & 기본 파일 존재 체크
    if not (f1 and f2):
        st.error("file1, file2 두 개의 엑셀을 먼저 업로드해 주세요.")
        st.stop()
    if not FILE3_PATH.exists():
        st.error(f"⚠️ 3번 파일이 {FILE3_PATH} 에 없습니다. 먼저 data 폴더와 파일을 리포지터리에 넣어 주세요.")
        st.stop()

    # 2) Excel → DataFrame
    df1 = pd.read_excel(f1)
    df2 = pd.concat(pd.read_excel(f2, sheet_name=None).values(), ignore_index=True)
    df3 = pd.read_excel(FILE3_PATH)

    # 3) 컬럼 선택
    c1  = pick(FILE1_COL_CAND, df1)
    c2  = pick(FILE2_COL_CAND, df2)
    c3  = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 4) 제목 정제
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 5) 1차 매핑 (file1 → file2)
    map1 = df1.drop_duplicates("정제_콘텐츠명").set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 6) 2차 매핑 (file3 → file2)
    map3 = df3.drop_duplicates("정제_콘텐츠3명").set_index("정제_콘텐츠3명")[id3]
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 7) 매핑콘텐츠명 / 콘텐츠ID (둘이 같은 행은 제외)
    mask_pair   = df2["정제_상품명"] == df2["매핑결과"]
    base_pairs  = (
        df2.loc[mask_pair, ["정제_상품명", "최종_매핑결과"]]
           .query("`정제_상품명`.str.strip() != ''", engine="python")
           .drop_duplicates()
           .rename(columns={"정제_상품명":"매핑콘텐츠명","최종_매핑결과":"콘텐츠ID"})
    )
    base_pairs["매핑콘텐츠명"] = base_pairs["매핑콘텐츠명"].apply(clean_title)

    unique_pairs = base_pairs[base_pairs["매핑콘텐츠명"] != base_pairs["콘텐츠ID"]]
    pad = len(df2) - len(unique_pairs)
    df2["매핑콘텐츠명"] = list(unique_pairs["매핑콘텐츠명"]) + [""] * pad
    df2["콘텐츠ID"]     = list(unique_pairs["콘텐츠ID"])     + [""] * pad

    # 8) file1 정보 붙이기
    info   = df1[[c1,"정제_콘텐츠명","판매채널콘텐츠ID"]].rename(columns={
              c1:"file1_콘텐츠명","정제_콘텐츠명":"file1_정제_콘텐츠명",
              "판매채널콘텐츠ID":"file1_판매채널콘텐츠ID"})
    result = pd.concat([df2, info], axis=1)

    # 9) 열 순서 재배치
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    result = result[front + [c for c in result.columns if c not in front]]

    # 10) 필요 없는 열 제거
    result.drop(
        columns=["최종_정렬된_매핑되지않은_상품명", "최종_매핑되지않은_상품명"],
        inplace=True, errors="ignore"
    )

    # 11) 엑셀 저장 & 헤더 색상
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

        wb, ws = writer.book, writer.sheets["매핑결과"]
        fmt_y  = wb.add_format({"bg_color":"#FFFFCC","bold":True,"border":1})

        for col_idx, col_name in enumerate(result.columns):
            if col_name in {"매핑콘텐츠명", "콘텐츠ID"}:
                ws.write(0, col_idx, col_name, fmt_y)

    # 12) 다운로드 버튼
    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        buf.getvalue(),
        file_name=save_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# app.py  ─────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd, re, io
import openpyxl, xlsxwriter          # requirements.txt 에도 명시!

# ────────── 후보 컬럼명 ──────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = ["컨텐츠","타이틀","작품명","도서명","작품 제목",
                  "상품명","이용상품명","상품 제목","ProductName","Title","제목"]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ────────── 유틸 ──────────
def pick(cands, df):
    for c in cands:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {cands}")

def clean_title(txt: str) -> str:
    t = str(txt)
    # ①  「 … 제 12권 」 류 제거
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)
    # ② 특수 치환
    for k,v in {"Un-holyNight":"UnholyNight","?":"","~":"","-":"","_":""," ,":""}.items():
        t = t.replace(k,v)
    # ③ 괄호·대괄호 제거
    t = re.sub(r"\([^)]*\)|\[[^\]]*\]", "", t)
    # ④ 숫자+권/화/부/회 제거
    t = re.sub(r"\d+[권화부회]", "", t)
    # ⑤ 키워드 제거
    for kw in ["개정판 l","개정판","외전","무삭제본","무삭제판","합본",
               "단행본","시즌","세트","연재","특별","최종화","완결",
               "2부","무삭제","완전판","세개정판","19세개정판"]:
        t = t.replace(kw,"")
    # ⑥ 나머지 숫자·기호 정리
    t = re.sub(r"\d+","",t).rstrip('.')
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}$begin:math:display$$end:math:display$$begin:math:text$$end:math:text$]","",t)
    t = re.sub(r"특별$","",t)
    return t.replace(" ","").strip()

# ────────── UI ──────────
st.title("📁 콘텐츠 매핑 도구 (웹)")

f1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
f2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
f3 = st.file_uploader("③ S2 콘텐츠 전체 (file3)", type="xlsx")

if st.button("🟢 매핑 실행"):

    if not (f1 and f2 and f3):
        st.error("3개의 엑셀 파일을 모두 업로드해 주세요.")
        st.stop()

    # ── 1) Excel → DataFrame ─────────────────────────────
    df1 = pd.read_excel(f1)
    df2 = pd.read_excel(f2, sheet_name=None)           # file2 는 다중 시트 허용
    df2 = pd.concat(df2.values(), ignore_index=True)
    df3 = pd.read_excel(f3)

    # ── 2) 컬럼 탐색 ──────────────────────────────────────
    c1 = pick(FILE1_COL_CAND, df1)
    c2 = pick(FILE2_COL_CAND, df2)
    c3 = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND , df3)

    # ── 3) 제목 정제 ────────────────────────────────────
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # ── 4) 1차 매핑 (file1 → file2) ─────────────────────
    map1 = df1.drop_duplicates("정제_콘텐츠명") \
              .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # ── 5) 2차 매핑 (file3 → file2) ─────────────────────
    map3 = df3.drop_duplicates("정제_콘텐츠3명") \
              .set_index("정제_콘텐츠3명")[id3]
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # ── 6) “매핑콘텐츠명 / 콘텐츠ID” 열 생성 ─────────────
    cond = (df2["정제_상품명"] == df2["매핑결과"]) &  \
           (df2["정제_상품명"] != df2["최종_매핑결과"])  # 숫자로 치환된 경우
    df2.loc[cond, "매핑콘텐츠명"] = df2.loc[cond, "정제_상품명"]
    df2.loc[cond, "콘텐츠ID"]   = df2.loc[cond, "최종_매핑결과"]

    # ── 7) “최종_정렬된_매핑되지않은_상품명” 열 ──────────
    no1 = df2.loc[df2["정제_상품명"] == df2["매핑결과"], "정제_상품명"]
    final_unmatch = no1[~no1.isin(map3.index)].drop_duplicates()
    df2["최종_정렬된_매핑되지않은_상품명"] = \
        sorted(final_unmatch) + [""]*(len(df2)-len(final_unmatch))

    # ── 8) file1 정보 + ‘미사용 콘텐츠’ 열 두 개 붙이기 ──
    info = df1[[c1,"정제_콘텐츠명","판매채널콘텐츠ID"]].rename(columns={
        c1:"file1_콘텐츠명","정제_콘텐츠명":"file1_정제_콘텐츠명",
        "판매채널콘텐츠ID":"file1_판매채널콘텐츠ID"
    })
    result = pd.concat([df2, info], axis=1)

    # file1 에서 사용 안 된 콘텐츠
    unused = df1[~df1["정제_콘텐츠명"].isin(df2["정제_상품명"])]
    unused = unused[["정제_콘텐츠명","판매채널콘텐츠ID"]] \
             .rename(columns={
                 "정제_콘텐츠명":"미매핑_정제_콘텐츠명",
                 "판매채널콘텐츠ID":"미매핑_콘텐츠ID"
             }).reset_index(drop=True)

    # 행 수 맞춰서 오른쪽 열 2개로 붙이기
    pad_len = len(result) - len(unused)
    if pad_len > 0:
        unused = pd.concat([unused, pd.DataFrame({"미매핑_정제_콘텐츠명":[""]*pad_len,
                                                  "미매핑_콘텐츠ID":[""]*pad_len})],
                           ignore_index=True)
    result = pd.concat([result, unused], axis=1)

    # ── 9) 결과 저장 & 다운로드 ──────────────────────────
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        out.getvalue(),
        file_name="mapping_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

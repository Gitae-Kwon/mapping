# app.py ───────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import re, io
import openpyxl, xlsxwriter      # requirements.txt 에 이미 포함

# ─── 후보 컬럼 ─────────────────────────────────────────────────────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = ["컨텐츠","타이틀","작품명","도서명","작품 제목",
                  "상품명","이용상품명","상품 제목","ProductName","Title","제목"]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID","콘텐츠ID","ID","ContentID"]

# ─── 공통 유틸 ─────────────────────────────────────────────────────────
def pick(cands, df):
    for c in cands:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {cands}")

def clean_title(txt:str) -> str:
    t = str(txt)
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)
    for k,v in {"Un-holyNight":"UnholyNight","?":"","~":"",",":"","-":"","_":""}.items():
        t = t.replace(k,v)
    t = re.sub(r"\([^)]*\)","",t);   t = re.sub(r"\[[^\]]*\]","",t)
    t = re.sub(r"\d+[권화부회]","",t)
    for kw in ["개정판 l","개정판","외전","무삭제본","무삭제판","합본",
               "단행본","시즌","세트","연재","특별","최종화","완결",
               "2부","무삭제","완전판","세개정판","19세개정판"]:
        t = t.replace(kw,"")
    t = re.sub(r"\d+","",t).rstrip('.')
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}()]", "", t)
    t = re.sub(r"특별$", "", t)
    return t.replace(" ","").strip()

# ─── Streamlit UI ─────────────────────────────────────────────────────
st.title("📁 콘텐츠 매핑 도구 (웹버전)")

file1 = st.file_uploader("① S2 채널 전체 (file1)", type="xlsx")
file2 = st.file_uploader("② 플랫폼 제공 정산서 (file2)", type="xlsx")
file3 = st.file_uploader("③ S2 콘텐츠 전체 (file3)", type="xlsx")

if st.button("🟢 매핑 실행"):

    if not (file1 and file2 and file3):
        st.error("3개의 엑셀 파일을 모두 업로드해 주세요.")
        st.stop()

    # 1) Excel → DataFrame
    df1 = pd.read_excel(file1)
    df2 = pd.concat(pd.read_excel(file2, sheet_name=None).values(), ignore_index=True)
    df3 = pd.read_excel(file3)

    # 2) 제목/ID 컬럼 선택
    c1 = pick(FILE1_COL_CAND, df1)
    c2 = pick(FILE2_COL_CAND, df2)
    c3 = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND, df3)

    # 3) 제목 정제
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 4) 1차 매핑 (file1 → file2)
    map1 = df1.drop_duplicates("정제_콘텐츠명").set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 5) 1차 미매핑
    no1 = df2.loc[df2["정제_상품명"] == df2["매핑결과"], "정제_상품명"]

    # 6) 2차 매핑 (file3 → file2)
    map3 = df3.drop_duplicates("정제_콘텐츠3명").set_index("정제_콘텐츠3명")[id3]
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 7) 최종 미매핑 & 정렬
    final_unmatch = no1[~no1.isin(map3.index)].drop_duplicates()
    df2["최종_정렬된_매핑되지않은_상품명"] = (
        sorted(final_unmatch) + [""] * (len(df2) - len(final_unmatch))
    )
    df2["최종_매핑되지않은_상품명"] = df2["정제_상품명"].where(
        df2["정제_상품명"].isin(final_unmatch), ""
    )
    # ── ✅ 2-A)  ‘매핑 성공한 행’ → 매핑콘텐츠명 / 콘텐츠ID 컬럼 채우기
    #   · 정제_상품명 ≠ 최종_매핑결과  → 매핑이 일어난 행
    success_mask = df2["정제_상품명"] != df2["최종_매핑결과"]

    df2.loc[success_mask, "매핑콘텐츠명"] = df2.loc[success_mask, "정제_상품명"]
    df2.loc[success_mask, "콘텐츠ID"]   = df2.loc[success_mask, "최종_매핑결과"]

  # ── 7-B)  “매핑콘텐츠명 / 콘텐츠ID” 컬럼 만들기 ─────────────────────────
# ① 두 컬럼을 미리 빈 값으로 만들어 둠
df2["매핑콘텐츠명"] = ""
df2["콘텐츠ID"]   = ""

# ② ➊ ‘정제_상품명’ == ‘매핑결과’(→ 아직 매핑되지 않음) ➋ 값이 숫자가 아닌 경우만
unmapped_mask = (
    (df2["정제_상품명"] == df2["매핑결과"]) &
    (~df2["매핑결과"].astype(str).str.isnumeric())
)

# ③ 첫 번째 등장 행에만 값을 채우고, 이후 중복은 비움
first_only = ~df2.loc[unmapped_mask, "정제_상품명"].duplicated()

df2.loc[unmapped_mask & first_only, "매핑콘텐츠명"] = df2.loc[unmapped_mask & first_only, "정제_상품명"]
#   콘텐츠ID 는 아직 ID가 없으므로 공란(필요하면 다른 정보로 채우세요)

    # 8) file1 정보 붙이기
    info = df1[[c1, "정제_콘텐츠명", "판매채널콘텐츠ID"]].rename(columns={
        c1: "file1_콘텐츠명",
        "정제_콘텐츠명": "file1_정제_콘텐츠명",
        "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID",
    })
    result = pd.concat([df2, info], axis=1)

    # ★ 9) **file1_* 세 열을 맨 앞으로 이동 (추가된 두 줄)**
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    result = result[front + [c for c in result.columns if c not in front]]

    # 10) 결과를 엑셀로 메모리에 저장
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드하세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        buffer.getvalue(),
        file_name="mapping_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

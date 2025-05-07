# app.py ───────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import re, io, pathlib
import openpyxl, xlsxwriter
from datetime import datetime, date

# ── (고정) ③번 파일 경로 ─────────────────────────────────────────────
DATA_DIR   = pathlib.Path(__file__).parent / "data"
FILE3_PATH = DATA_DIR / "all_contents.xlsx"

# ── 후보 컬럼 ─────────────────────────────────────────────────────────
FILE1_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_COL_CAND = [
    "컨텐츠", "타이틀", "작품명", "도서명", "작품 제목",
    "상품명", "이용상품명", "상품 제목", "ProductName", "Title", "제목", "컨텐츠명", "콘텐츠명"
]
FILE3_COL_CAND = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CAND  = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ── 유틸 ──────────────────────────────────────────────────────────────
def pick(cands, df):
    for c in cands:
        if c in df.columns:
            return c
    raise ValueError(f"가능한 컬럼이 없습니다 ➜ {cands}")

def clean_title(txt) -> str:
    # 1) 진짜 날짜(datetime/date) 객체면 f"{월}월{일}일" 로
    if isinstance(txt, (datetime, date)):
        return f"{txt.month}월{txt.day}일"

    # 2) 이미 "7월24일" 처럼 포맷된 문자열이면 그대로
    t = str(txt).strip()
    if re.fullmatch(r"\d{1,2}월\d{1,2}일", t):
        return t

    # 3) 그 외 기존 정제 로직
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)
    for k, v in {
        "Un-holyNight": "UnholyNight", "?": "", "~": "", ",": "", "-": "", "_": ""
    }.items():
        t = t.replace(k, v)
    t = re.sub(r"\([^)]*\)|\[[^\]]*\]", "", t)
    t = re.sub(r"\d+[권화부회]", "", t)
    for kw in [
        "개정판 l","개정판","외전","무삭제본","무삭제판","합본",
        "단행본","시즌","세트","연재","특별","최종화","완결",
        "2부","무삭제","완전판","세개정판","19세개정판"
    ]:
        t = t.replace(kw, "")
    t = re.sub(r"\d+", "", t).rstrip(".")
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､{}()]", "", t)
    t = t.replace("[", "").replace("]", "")
    t = re.sub(r"특별$", "", t)
    return t.replace(" ", "").strip()

# ── UI ────────────────────────────────────────────────────────────────
st.title("📁 판매채널 및 콘텐츠마스터ID 매핑")

# ① S2-판매채널 콘텐츠리스트 업로드
f1 = st.file_uploader(
    "① S2-판매채널 콘텐츠리스트  "
    "( https://kiss.kld.kr/mst/sch/schn-ctns-search )에서 판매채널을 선택 조회 후 엑셀 다운로드  \n"
    "※ S2에서 다운로드된한 파일은 ‘열기’ → ‘다른 이름으로 저장’ 후 업로드해 주세요.(바로 사용시 오류발생)",
    type="xlsx",
)

# ② 플랫폼별 정산서 업로드
f2 = st.file_uploader(
    "② 플랫폼별 정산서 (판매채널에서 제공한 정산서)",
    type="xlsx",
)

# ③ A/B 법인 선택 (고정된 data 폴더 내 파일 사용)
choice3 = st.selectbox(
    "③ 콘텐츠마스터 매핑 법인을 선택해주세요",
    ("키다리스튜디오", "레진KR", "레진JP"),
    help="선택한 법인을 기준으로 IPS 콘텐츠마스터 ID와 매핑 합니다."
)

# 선택에 따라 사용할 3번 파일 경로 결정
if choice3 == "키다리스튜디오":
    file3_path = DATA_DIR / "kidari_contents.xlsx"
elif choice3 == "레진KR":
    file3_path = DATA_DIR / "lezhin_contents.xlsx"
else:  # 레진 JP
    file3_path = DATA_DIR / "lezhinjp_contents.xlsx"

# ④ 저장 파일명 기본값: 업로드한 f2 파일명(stem) + '매핑'
from pathlib import Path

if f2 is not None:
    default_name = Path(f2.name).stem + "_매핑"
else:
    default_name = "mapping_result"

# 사용자가 저장 파일명을 변경할 수 있고, 내부에서 .xlsx를 붙입니다
save_name = st.text_input(
    "💾 저장 파일명(확장자 제외)",
    value=default_name
) + ".xlsx"

# ── 매핑 실행 버튼 ─────────────────────────────────────────────────────
if st.button("🟢 매핑 실행"):
    # 업로드 확인
    if not (f1 and f2):
        st.error("file1, file2를 모두 업로드해 주세요.")
        st.stop()
    # 3번 파일 확인
    if not file3_path.exists():
        st.error(f"선택된 3번 파일이 `{file3_path}` 에 없습니다.")
        st.stop()

    # 1) Excel → DataFrame
    df1 = pd.read_excel(f1)
    df2 = pd.concat(pd.read_excel(f2, sheet_name=None).values(), ignore_index=True)
    df3 = pd.read_excel(file3_path)  # A/B 중 사용자가 선택한 파일 읽기

    # 3) 컬럼 선택 -----------------------------------------------------
    c1  = pick(FILE1_COL_CAND, df1)
    c2  = pick(FILE2_COL_CAND, df2)
    c3  = pick(FILE3_COL_CAND, df3)
    id3 = pick(FILE3_ID_CAND,  df3)

    # 4) 제목 정제 -----------------------------------------------------
    df1["정제_콘텐츠명"]  = df1[c1].apply(clean_title)
    df2["정제_상품명"]    = df2[c2].apply(clean_title)
    df3["정제_콘텐츠3명"] = df3[c3].apply(clean_title)

    # 5) 1차 매핑 -----------------------------------------------------
    map1 = (
        df1.drop_duplicates("정제_콘텐츠명")
           .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
    )
    df2["매핑결과"] = df2["정제_상품명"].map(map1).fillna(df2["정제_상품명"])

    # 6) 2차 매핑 -----------------------------------------------------
    map3 = (
        df3.drop_duplicates("정제_콘텐츠3명")
           .set_index("정제_콘텐츠3명")[id3]
    )
    df2["최종_매핑결과"] = df2["정제_상품명"].map(map3).fillna(df2["매핑결과"])

    # 7) 매핑콘텐츠명 / 콘텐츠ID --------------------------------------
    mask_pair  = df2["정제_상품명"] == df2["매핑결과"]
    base_pairs = (
        df2.loc[mask_pair, ["정제_상품명", "최종_매핑결과"]]
           .query("`정제_상품명`.str.strip() != ''", engine="python")
           .drop_duplicates()
           .rename(columns={
               "정제_상품명": "매핑콘텐츠명",
               "최종_매핑결과": "콘텐츠ID"
           })
    )
    base_pairs["매핑콘텐츠명"] = base_pairs["매핑콘텐츠명"].apply(clean_title)

    dup_mask     = base_pairs["매핑콘텐츠명"] == base_pairs["콘텐츠ID"]
    pairs_unique = (
        base_pairs.loc[~dup_mask]
                  .sort_values("매핑콘텐츠명")
                  .reset_index(drop=True)
    )
    pairs_same   = (
        base_pairs.loc[dup_mask]
                  .sort_values("매핑콘텐츠명")
                  .reset_index(drop=True)
    )

    pad_u = len(df2) - len(pairs_unique)
    df2["매핑콘텐츠명"] = list(pairs_unique["매핑콘텐츠명"]) + [""] * pad_u
    df2["콘텐츠ID"]     = list(pairs_unique["콘텐츠ID"])     + [""] * pad_u

    pad_s = len(df2) - len(pairs_same)
    df2["동일_매핑콘텐츠명"] = list(pairs_same["매핑콘텐츠명"]) + [""] * pad_s
    df2["동일_콘텐츠ID"]     = list(pairs_same["콘텐츠ID"])     + [""] * pad_s

    # 8) 최종 미매핑 ---------------------------------------------------
    used_titles   = set(base_pairs["매핑콘텐츠명"])
    final_unmatch = (
        df2.loc[mask_pair, "정제_상품명"]
           .drop_duplicates()
           .pipe(lambda s: s[~s.isin(map3.index)])
           .pipe(lambda s: s[~s.isin(used_titles)])
    )
    df2["최종_정렬된_매핑되지않은_상품명"] = (
        sorted(final_unmatch) + [""] * (len(df2) - len(final_unmatch))
    )
    df2["최종_매핑되지않은_상품명"] = df2["정제_상품명"].where(
        df2["정제_상품명"].isin(final_unmatch), ""
    )

    # 9) file1 정보 ---------------------------------------------------
    info = (
        df1[[c1, "정제_콘텐츠명", "판매채널콘텐츠ID"]]
           .rename(columns={
               c1: "file1_콘텐츠명",
               "정제_콘텐츠명": "file1_정제_콘텐츠명",
               "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID"
           })
    )
    result = pd.concat([df2, info], axis=1)
    insert_at = result.columns.get_loc("정제_상품명") + 1
    result.insert(insert_at, "정산서_콘텐츠명", result[c2])

    # 10) 열 순서 재배치 ---------------------------------------------
    front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
    cols  = list(result.columns)
    idx   = cols.index("콘텐츠ID") + 1
    cols.remove("동일_매핑콘텐츠명")
    cols.remove("동일_콘텐츠ID")
    cols[idx:idx] = ["동일_매핑콘텐츠명", "동일_콘텐츠ID"]
    result = result[front + [c for c in cols if c not in front]]

    # 11) 필요 없는 열 제거 ------------------------------------------
    result.drop(
        columns=[
            "동일_콘텐츠ID",
            "최종_정렬된_매핑되지않은_상품명",
            "최종_매핑되지않은_상품명"
        ],
        inplace=True,
        errors="ignore"
    )

   # ★ 12-a) 컬럼명 변경
result.rename(columns={
    "매핑콘텐츠명":           "매핑_콘텐츠마스터명",
    "콘텐츠ID":             "매핑_콘텐츠마스터ID",
    "동일_매핑콘텐츠명":     "미매핑_콘텐츠마스터명",
    "file1_콘텐츠명":           "S2_콘텐츠명",
    "file1_정제_콘텐츠명":      "S2_정제콘텐츠명",
    "file1_판매채널콘텐츠ID":   "S2_판매채널콘텐츠ID",
    "매핑결과":               "매핑_판매채널콘텐츠ID",
    "최종_매핑결과":            "매핑_콘텐츠마스터ID",
}, inplace=True)

    
    # 12) 엑셀 저장 + 헤더 서식 + 숨김처리 ─────────────────────────────
buf = io.BytesIO()
    visible_cols = {            # ❖ 숨기지 않을 열
        "S2_콘텐츠명",
        "S2_정제콘텐츠명",
        "S2_판매채널콘텐츠ID",
        "정제_상품명",
        "매핑_판매채널콘텐츠ID",
        "매핑_콘텐츠마스터ID",
        "매핑_콘텐츠마스터명",
        "미매핑_콘텐츠마스터명",
        "정산서_콘텐츠명",
    }
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="매핑결과", index=False)

        wb = writer.book
        ws = writer.sheets["매핑결과"]
        # ⑥ ★ 헤더 길이에 맞춰 열 너비 자동 조정
        for col_idx, col_name in enumerate(result.columns):
            width = len(col_name) + 1
            ws.set_column(col_idx, col_idx, width)    

        # 헤더 색상
        fmt_yellow = wb.add_format({"bg_color": "#FFFFCC", "bold": True, "border": 1})
        fmt_green  = wb.add_format({"bg_color": "#99FFCC", "bold": True, "border": 1})

        for col_idx, col_name in enumerate(result.columns):
            # 서식
            if col_name in {"매핑_콘텐츠마스터명", "매핑_콘텐츠마스터ID"}:
                ws.write(0, col_idx, col_name, fmt_yellow)
            elif col_name == "미매핑_콘텐츠마스터명":
                ws.write(0, col_idx, col_name, fmt_green)

            # 숨김
            if col_name not in visible_cols:
                ws.set_column(col_idx, col_idx, None, None, {"hidden": True})

    # 13) 다운로드 ----------------------------------------------------
    st.success("✅ 매핑 완료! 아래 버튼으로 다운로드해주세요.")
    st.download_button(
        "📥 결과 엑셀 다운로드",
        buf.getvalue(),
        file_name=save_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import xlsxwriter  # XlsxWriter 엔진 사용

# ─── openpyxl CellStyle 패치 ──────────────────────────────────────────
try:
    import openpyxl.styles.cell as _cellmodule
    _orig_CellStyle = _cellmodule.CellStyle

    def _patched_init(self, *args, **kwargs):
        kwargs.pop('count', None)          # openpyxl 3.1 이상에서만 나오는 인자 제거
        return _orig_CellStyle.__init__(self, *args, **kwargs)

    _cellmodule.CellStyle.__init__ = _patched_init
except ImportError:
    pass
# ---------------------------------------------------------------------

# 후보 컬럼명
FILE1_TITLE_CANDIDATES = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE2_TITLE_CANDIDATES = ["컨텐츠", "타이틀", "작품명", "도서명", "작품 제목", "상품명", "이용상품명", "제목", "상품 제목", "ProductName", "Title", "제목"]
FILE3_TITLE_CANDIDATES = ["콘텐츠명", "콘텐츠 제목", "Title", "ContentName", "제목"]
FILE3_ID_CANDIDATES    = ["판매채널콘텐츠ID", "콘텐츠ID", "ID", "ContentID"]

# ─── 유틸 함수 ────────────────────────────────────────────────────────
def get_title_column(df: pd.DataFrame, candidates) -> str:
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"제목 컬럼을 찾을 수 없습니다. 후보: {candidates}")

def get_id_column(df: pd.DataFrame, candidates) -> str:
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"ID 컬럼을 찾을 수 없습니다. 후보: {candidates}")

def clean_title(raw: str) -> str:
    t = str(raw)

    # “ 제숫자권/화” 패턴 제거
    t = re.sub(r"\s*제\s*\d+[권화]", "", t)

    # 특수 치환 ──❶ 이 부분만 수정
    special_cases = {
        "Un-holyNight": "UnholyNight",
        "?":  "",
        "~":  "",
        ",":  "",
        "-":  "",
        "_":  ""
    }
    for k, v in special_cases.items():
        t = t.replace(k, v)

    # 괄호 내용 제거
    t = re.sub(r"\([^)]*\)", "", t)
    t = re.sub(r"\[[^\]]*\]", "", t)

    # 숫자+권/화/부/회 제거
    t = re.sub(r"\d+[권화부회]", "", t)

    # 불필요 키워드 제거
    keywords = sorted(
        ["개정판 l", "개정판", "외전", "무삭제본", "무삭제판", "합본",
         "단행본", "시즌", "세트", "연재", "특별", "최종화", "완결",
         "2부", "무삭제", "완전판", "세개정판", "19세개정판"],
        key=lambda x: -len(x)
    )
    for w in keywords:
        t = t.replace(w, "")

    # 모든 숫자 제거
    t = re.sub(r"\d+", "", t)

    # 끝점·특수문자·마지막 '특별' 정리
    t = t.rstrip('.')
    t = re.sub(r"[\.~\-–—!@#$%^&*_=+\\|/:;\"'’`<>?，｡､$begin:math:display$$end:math:display$$begin:math:text$$end:math:text$\{\}]", "", t)
    t = re.sub(r"특별$", "", t)

    return t.replace(" ", "").strip()

# ─── 파일 다이얼로그 ────────────────────────────────────────────────
def browse_file(entry: tk.Entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# ─── 매핑 실행 ───────────────────────────────────────────────────────
def run_mapping():
    try:
        # 0) 경로 읽기
        path1, path2, path3 = (e.get() for e in entries)

        # 1) 엑셀 로드
        df1 = pd.read_excel(path1)
        df2 = pd.read_excel(path2)
        df3 = pd.read_excel(path3)

        # 2) 컬럼 찾기
        col1 = get_title_column(df1, FILE1_TITLE_CANDIDATES)
        col2 = get_title_column(df2, FILE2_TITLE_CANDIDATES)
        col3 = get_title_column(df3, FILE3_TITLE_CANDIDATES)
        id3  = get_id_column(df3,  FILE3_ID_CANDIDATES)

        # 3) 제목 정제
        df1["정제_콘텐츠명"]  = df1[col1].apply(clean_title)
        df2["정제_상품명"]    = df2[col2].apply(clean_title)
        df3["정제_콘텐츠3명"] = df3[col3].apply(clean_title)

        # 4) 1차 매핑 (file1 → file2)
        mapping1 = (
            df1.drop_duplicates("정제_콘텐츠명")
               .set_index("정제_콘텐츠명")["판매채널콘텐츠ID"]
        )
        df2["매핑결과"] = df2["정제_상품명"].map(mapping1).fillna(df2["정제_상품명"])

        # 5) 1차 미매핑
        unmatched1 = df2.loc[df2["정제_상품명"] == df2["매핑결과"], "정제_상품명"]

        # 6) 2차 매핑 (file3 → file2)
        mapping3 = (
            df3.drop_duplicates("정제_콘텐츠3명")
               .set_index("정제_콘텐츠3명")[id3]
        )
        df2["최종_매핑결과"] = (
            df2["정제_상품명"].map(mapping3).fillna(df2["매핑결과"])
        )

        # 7) 최종 미매핑 리스트 & 정렬
        final_unmatched = (
            unmatched1[~unmatched1.isin(mapping3.index)]
            .drop_duplicates()
        )
        sorted_unmatched = sorted(final_unmatched.tolist())
        pad_len = len(df2) - len(sorted_unmatched)
        df2["최종_정렬된_매핑되지않은_상품명"] = sorted_unmatched + [""] * pad_len
        df2["최종_매핑되지않은_상품명"] = df2["정제_상품명"].apply(
            lambda x: x if x in final_unmatched.values else ""
        )

        # 8) file1 정보 붙이기
        df1_info = df1[[col1, "정제_콘텐츠명", "판매채널콘텐츠ID"]].rename(columns={
            col1: "file1_콘텐츠명",
            "정제_콘텐츠명": "file1_정제_콘텐츠명",
            "판매채널콘텐츠ID": "file1_판매채널콘텐츠ID",
        })
        result = pd.concat([df2, df1_info], axis=1)
        front = ["file1_콘텐츠명", "file1_정제_콘텐츠명", "file1_판매채널콘텐츠ID"]
        result = result[front + [c for c in result.columns if c not in front]]

        # 9) 저장
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="mapping_result.xlsx",
        )
        if not save_path:
            return

        with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
            result.to_excel(writer, sheet_name="매핑결과", index=False)

        messagebox.showinfo("완료", f"✅ 매핑 완료! 파일 저장됨: {save_path}")

    except Exception as err:
        messagebox.showerror("에러 발생", f"오류가 발생했습니다:\n{err}")

# ─── GUI ────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("📁 콘텐츠 매핑 도구")

labels = [
    "📂 파일1 (S2 채널전체):",
    "📂 파일2 (플랫폼제공 정산서):",
    "📂 파일3 (S2 콘텐츠 전체):",
]
entries = []

for i, text in enumerate(labels):
    tk.Label(root, text=text).grid(row=i, column=0, sticky="e")

    ent = tk.Entry(root, width=50)
    ent.grid(row=i, column=1)
    entries.append(ent)

    tk.Button(
        root,
        text="찾아보기",
        command=lambda e=ent: browse_file(e)   # 각 버튼이 해당 Entry를 캡처
    ).grid(row=i, column=2)

tk.Button(
    root,
    text="🟢 매핑 실행",
    width=30,
    bg="#4CAF50",
    fg="white",
    command=run_mapping,
).grid(row=3, column=0, columnspan=3, pady=15)

root.mainloop()

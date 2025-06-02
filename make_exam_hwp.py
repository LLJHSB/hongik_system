import pandas as pd
import win32com.client
import os
import random
import math

# === 설정 ===
엑셀파일 = "test_id.xlsx"   # 파일명
출력파일명 = "맞춤형_시험지"
문제수 = 5
사용자레벨 = "중상"  # 선택 가능: 최상, 상, 중상, 중, 중하, 하

# === 난이도 범위 (정답률 기준) ===
def get_level(rate):
    if rate < 30:
        return '최상'
    elif rate < 50:
        return '상'
    elif rate < 60:
        return '중상'
    elif rate < 80:
        return '중'
    elif rate < 90:
        return '중하'
    else:
        return '하'

# === 난이도별 비율 설정 ===
level_ratio = {
    "중상": {"상": 0.3, "중상": 0.4, "중": 0.3},
    "중":   {"중상": 0.2, "중": 0.5, "중하": 0.3},
    "하":   {"중하": 0.4, "하": 0.6},
    "상":   {"최상": 0.2, "상": 0.5, "중상": 0.3}
}

# === 데이터 불러오기 ===
df = pd.read_excel(엑셀파일)
df['난이도'] = df['정답률'].apply(get_level)

# === 문제 선택 ===
selected_problems = pd.DataFrame()
ratios = level_ratio.get(사용자레벨, {})
for lv, ratio in ratios.items():
    sub = df[(df['유형'] == '문제') & (df['난이도'] == lv)]
    count = math.ceil(문제수 * ratio)
    if len(sub) > 0:
        selected = sub.sample(min(count, len(sub)), replace=False)
        selected_problems = pd.concat([selected_problems, selected], ignore_index=True)

# === 첫 문제 기준으로 지문 찾기 ===
first_passage_id = selected_problems.iloc[0]['지문id']
passage_row = df[(df['유형'] == '지문') & (df['지문id'] == first_passage_id)].iloc[0]

# === 한글 실행 ===
hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# === 최종 문서 생성 ===
hwp.XHwpDocuments.Add(0)  # 새 문서

# ✅ 지문 붙여넣기
hwp.Open(os.path.abspath(passage_row['파일명']))
hwp.HAction.Run("SelectAll")
hwp.HAction.Run("Copy")
hwp.XHwpDocuments.Add(0)  # 시험지 문서 새로
hwp.HAction.Run("Paste")

# ✅ 문제들 붙여넣기
for _, row in selected_problems.iterrows():
    hwp.Open(os.path.abspath(row['파일명']))
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Copy")
    hwp.HAction.Run("MoveDocEnd")  # 커서 맨 끝으로
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("Paste")

# ✅ 저장
hwp_path = os.path.abspath(f"{출력파일명}.hwp")
pdf_path = os.path.abspath(f"{출력파일명}.pdf")
hwp.SaveAs(hwp_path)

# ✅ PDF 저장
hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileSaveAsPdf.HSet)
hwp.HParameterSet.HFileSaveAsPdf.filename = pdf_path
hwp.HParameterSet.HFileSaveAsPdf.Compatibility = False
hwp.HParameterSet.HFileSaveAsPdf.PDFSecurity = False
hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileSaveAsPdf.HSet)

print("✅ 시험지 생성 완료:", pdf_path)

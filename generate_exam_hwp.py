# 맞춤형 수능 국어 시험지 생성기
# 작성자: 졸업프로젝트 팀
# 사용법: 같은 폴더에 문제/지문 .hwp 파일들과 test_id.xlsx 파일을 두고 실행
# 필요: 한글(HWP) 설치 + Windows + pywin32

import pandas as pd
import win32com.client
import os
import random
import math

# === 설정 ===
엑셀파일 = "test_id.xlsx"   # 문제/지문 메타 정보
출력파일명 = "맞춤형_시험지"  # 출력 파일명 (hwp/pdf)
문제수 = 5
사용자레벨 = "중상"  # 선택 가능: 최상, 상, 중상, 중, 중하, 하

# === 정답률 → 난이도 변환 함수 ===
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

# === 사용자 레벨별 문제 난이도 비율 설정 ===
level_ratio = {
    "중상": {"상": 0.3, "중상": 0.4, "중": 0.3},
    "중":   {"중상": 0.2, "중": 0.5, "중하": 0.3},
    "하":   {"중하": 0.4, "하": 0.6},
    "상":   {"최상": 0.2, "상": 0.5, "중상": 0.3}
}

# === 1. 데이터 불러오기 ===
df = pd.read_excel(엑셀파일)
df['난이도'] = df['정답률'].apply(get_level)

# === 2. 난이도별 비율로 문제 선택 ===
selected_problems = pd.DataFrame()
ratios = level_ratio.get(사용자레벨, {})
for lv, ratio in ratios.items():
    sub = df[(df['유형'] == '문제') & (df['난이도'] == lv)]
    count = math.ceil(문제수 * ratio)
    if len(sub) > 0:
        selected = sub.sample(min(count, len(sub)), replace=False)
        selected_problems = pd.concat([selected_problems, selected], ignore_index=True)

# === 3. 첫 문제 기준 지문 선택 ===
first_passage_id = selected_problems.iloc[0]['지문id']
passage_row = df[(df['유형'] == '지문') & (df['지문id'] == first_passage_id)].iloc[0]

# === 4. 한글 실행 및 새 문서 생성 ===
hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.XHwpDocuments.Add(0)
main_doc = hwp.XHwpDocuments.Item(0)

# === 5. 지문 복사 → 시험지 문서에 붙여넣기 ===
hwp.Open(os.path.abspath(passage_row['파일명']))
hwp.HAction.Run("SelectAll")
hwp.HAction.Run("Copy")
main_doc.Activate()
hwp.HAction.Run("Paste")

# === 6. 선택된 문제들 순서대로 복사/붙여넣기 ===
for _, row in selected_problems.iterrows():
    hwp.Open(os.path.abspath(row['파일명']))
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Copy")
    main_doc.Activate()
    hwp.HAction.Run("MoveDocEnd")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("Paste")

# === 7. .hwp로 저장 ===
main_doc.Activate()
hwp.SaveAs(os.path.abspath(f"{출력파일명}.hwp"))

# === 8. .pdf로 저장 ===
hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileSaveAsPdf.HSet)
hwp.HParameterSet.HFileSaveAsPdf.filename = os.path.abspath(f"{출력파일명}.pdf")
hwp.HParameterSet.HFileSaveAsPdf.Compatibility = False
hwp.HParameterSet.HFileSaveAsPdf.PDFSecurity = False
hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileSaveAsPdf.HSet)

print("✅ 시험지 PDF 생성 완료:", f"{출력파일명}.pdf")

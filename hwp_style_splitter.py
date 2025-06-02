import os
import win32com.client as win32

# 파일 경로 지정
src_path = r"C:\Users\사용자이름\Desktop\문제지.hwp"
dst_folder = r"C:\Users\사용자이름\Desktop\분리된문제들"

# 저장 폴더가 없다면 생성
os.makedirs(dst_folder, exist_ok=True)

# 한글 실행
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# 파일 열기
hwp.Open(src_path)

# 스타일 이름
TARGET_STYLE = "문제번호"  # 예: 문제번호 스타일 이름

# 문서 시작으로 이동
hwp.MoveDocBegin()

# 문제 카운터
q_num = 1

# 반복 탐색
while True:
    style = hwp.GetCurFieldName()

    if style == TARGET_STYLE:
        # 스타일이 발견되면 문제 시작 위치 저장
        hwp.Run("SelectCtrlFront")
        hwp.Run("Copy")

        # 새 문서 만들고 붙여넣기
        hwp.Create("Blank", "HWP")
        hwp.Run("Paste")

        # 새 파일로 저장
        save_path = os.path.join(dst_folder, f"문제_{q_num}.hwp")
        hwp.SaveAs(save_path)
        hwp.Quit()

        # 기존 문서 다시 열기
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(src_path)
        hwp.MoveDocBegin()

        # 문제 번호만큼 아래로 이동
        for _ in range(q_num):
            hwp.Run("MoveNextPara")

        q_num += 1

    # 다음 문단으로 이동, 더 이상 없으면 종료
    if not hwp.Run("MoveNextPara"):
        break

print(f"{q_num - 1}개의 문제 문서를 저장했습니다.")

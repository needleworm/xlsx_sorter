"""
Author : Byunghyun Ban
Last modification : 2020.12.24.
bhban@kakao.com
https://github.com/needleworm/xlsx_sorter
"""
import time
import sys
import os
import pyexcel as px

print("Process Start")
start_time = time.time()

# 예시로 사용할 템플릿 엑셀 파일 이름
template = sys.argv[1]

# 분석 대상이 되는 엑셀 파일들이 들어있는 폴더
directory = sys.argv[2]

# 프로그램의 작동 모드 (delete, report, separate)
mode = sys.argv[3]

# 분석 대상이 되는 엑셀 파일의 목록
file_list = os.listdir(directory)

# 템플릿 파일을 읽어와 헤더를 분리합니다
HEADER = px.get_array(file_name=template)[0]

# 보고 모드일 때에는 보고서 파일을 생성합니다.
if mode == "report" or mode == "REPORT":
    # report.txt 파일을 새로이 생성합니다.
    report = open("report.txt", 'w')

# 분리 모드일 때에는 분리된 파일을 격리저장할 폴더를 만듭니다.
elif mode == "separate" or mode == "SEPARATE":
    import shutil
    # wrong_files 라는 폴더를 새로이 생성합니다.
    os.mkdir("wrong_files")

elif mode != "delete" and mode != "DELETE":
    print("Wrong Mode! (delete / report / separate)")
    exit(1)

# for문을 활용해 파일을 하나씩 불러 옵니다
for filename in file_list:
    # 엑셀 파일을 읽어옵니다.
    file = px.get_array(file_name=directory + "/" + filename)
    # 헤더를 분리합니다.
    header = file[0]
    # 헤더가 템플릿과 일치하는지 분석합니다.
    if header == HEADER:
        # 헤더가 템플릿과 일치하는 올바른 파일이라면
        # 아무것도 하지 않고 넘어갑니다.
        continue

    # 삭제 모드인 경우
    if mode in "DELETE delete":
        # 파일을 삭제합니다
        os.remove(directory + "/" + filename)
    # 보고 모드인 경우
    elif mode in "report REPORT":
        # 보고서에 파일 이름을 적습니다
        report.write(filename + "\n")
    # 분리 모드인 경우
    else:
        # 파일을 이동시킵니다.
        shutil.move(directory + "/" + filename, "wrong_files/" + filename)

end_time = time.time()
print("Process Done.")
print("The Job Took " + str(end_time - start_time) + " seconds.")

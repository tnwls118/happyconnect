from select import select
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import requests
import shutil
import datetime
import math
import time
import os
import os.path
import warnings
import traceback

# 특정 경고문 무시
warnings.filterwarnings(
    "ignore", message="Workbook contains no default style, apply openpyxl's default")

# 다운로드 폴더 경로 지정
files_Path = os.path.expanduser('~')+'\\Downloads'+'\\'
print("다운로드 경로:", files_Path)

# 크롬드라이버 경로 찾기


def chrome_driver_path():
    print('Chrome드라이브 경로 설정')
    time.sleep(0.5)
    # 크롬드라이버 설치된 경로로 수정
    return r"C:\Users\USER\Desktop\24\chromedriver.exe"


# 시간 산출
start = time.time()
math.factorial(100000)

search_Date_Len = int(input("\n검색 일자를 입력하시오(숫자+엔터/89일 초과 금지)\n"))

temp_Today = datetime.datetime.now()
temp_Start = temp_Today - datetime.timedelta(days=search_Date_Len)

# 검색 시작일
start_Year = temp_Start.year
start_Month = temp_Start.month
start_Day = temp_Start.day

# 검색 종료일
end_Year = temp_Today.year
end_Month = temp_Today.month
end_Day = temp_Today.day

homepage_Path = "https://happycommunity.happyconnect.co.kr/"

# 사용변수 List
driver_Happy = ""
file_Number = 6
loc_Address = ""
data_Origin = ""
rcp_Data = ""
data_Extract = ""

try:
    print("프로그램 시작")
    driver_Happy = webdriver.Chrome(chrome_driver_path())
    # 파일 숫자
    # 누구 디벨로퍼 접속시작
    driver_Happy.maximize_window()
    driver_Happy.get(homepage_Path)
    # 로그인 시작
    driver_Happy.find_element_by_id("user_id").send_keys("hc_kjw")
    driver_Happy.find_element_by_id("pwd").send_keys("!happyeco04@")
    driver_Happy.find_element_by_id("CO000_EVT001").click()
    driver_Happy.implicitly_wait(60)

    # ################################## 데이터 수집 시작 ############################################
    print("데이터 수집 시작")
    # # 메뉴 이동
    driver_Happy.find_element_by_xpath(
        """//*[@id="header-top"]/div[2]/div[1]/div/ul/li[7]/h2/a""").click()
    # # AI스피커 이동
    driver_Happy.find_element_by_xpath(
        """//*[@id="header-top"]/div[2]/div[1]/div/ul/li[7]/ul/li[1]/ul/li[2]/a""").click()
    driver_Happy.implicitly_wait(60)
    ####################### 시작일 선택 #####################
    # 시작 켈린더 클릭
    driver_Happy.find_element_by_xpath(
        """//*[@id="search_dtime_from"]""").click()
    # 년 선택
    cal_Selector = Select(driver_Happy.find_element_by_xpath(
        """//*[@id="ui-datepicker-div"]/div/div/select[1]"""))
    cal_Selector.select_by_visible_text('{}'.format(start_Year))
    # 월 선택
    cal_Selector = Select(driver_Happy.find_element_by_xpath(
        """//*[@id="ui-datepicker-div"]/div/div/select[2]"""))
    cal_Selector.select_by_visible_text('{}월'.format(start_Month))
    temp_table = driver_Happy.find_element_by_xpath(
        """/html/body/div[2]/table""")
    temp_tbody = temp_table.find_element_by_tag_name("tbody")
    start_tr_index = 0
    start_td_index = 0
    # 날짜 선택
    for tr in temp_tbody.find_elements_by_tag_name("tr"):
        start_tr_index += 1
        # td 인덱스 초기화
        start_td_index = 0
        for td in tr.find_elements_by_tag_name("td"):
            start_td_index += 1
            if td.text == '{}'.format(start_Day):
                driver_Happy.find_element_by_xpath(
                    """/html/body/div[2]/table/tbody/tr[{}]/td[{}]/a""".format(start_tr_index, start_td_index)).click()
                break
    time.sleep(2)
    #######################################################

    # 지역이동 시작
    loc_Address = pd.read_excel("심리상담지역.xlsx")
    대구분 = ""
    중구분 = ""
    소구분 = ""

    total_Loc = len(loc_Address['대구분'])
    print("총 지역수 : {}개".format(total_Loc))

    try:
        for i in range(len(loc_Address['대구분'])):
            print("현재 : {}/{}번째".format(i+1, total_Loc))

            file_Number += 1
            대구분 = loc_Address['대구분'][i]
            중구분 = loc_Address['중구분'][i]
            소구분 = loc_Address['소구분'][i]

            top_Select = Select(driver_Happy.find_element_by_xpath(
                """//*[@id="AN002001_FORM"]/div[1]/div[1]/ul[1]/li/select[1]"""))
            top_Select.select_by_visible_text(대구분)
            time.sleep(1)

            middle_Select = Select(driver_Happy.find_element_by_xpath(
                """//*[@id="AN002001_FORM"]/div[1]/div[1]/ul[1]/li/select[2]"""))
            middle_Select.select_by_visible_text(중구분)
            time.sleep(1)

            bottom_Select = Select(driver_Happy.find_element_by_xpath(
                """//*[@id="AN002001_FORM"]/div[1]/div[1]/ul[1]/li/select[3]"""))
            bottom_Select.select_by_visible_text(소구분)
            time.sleep(1)

            # 검색 버튼 클릭
            driver_Happy.find_element_by_xpath(
                """//*[@id="AN002001_SEARCH"]""").click()
            time.sleep(60)

            # 엑셀 다운로드 버튼 클릭
            driver_Happy.find_element_by_xpath(
                """//*[@id="AN002001_EVT016"]""").click()
            time.sleep(1)

            driver_Happy.find_element_by_xpath(
                """//*[@id="exelDownSystem"]""").click()
            time.sleep(1)

            driver_Happy.find_element_by_xpath(
                """//*[@id="CM001001_EVT001"]""").click()

            if 대구분 == "지자체":
                time.sleep(30)
            else:
                time.sleep(10)
    except:
        print("없는 지역입니다<{}-{}-{}>".format(대구분, 중구분, 소구분))
    stemp = time.time()
    print(f"데이터 수집 종료까지 {stemp - start:.1f} 초")
    ################################## 데이터 수집 끝 ##############################################

    time.sleep(60)

    ################################# 데이터 분류 시작 ############################################
    # 데이터 추출 Entry Point
    # 파일 변환 시작

    # 데이터 분류 시작
    print("데이터 분류 시작")
    data_Extract = pd.DataFrame(columns=['사용일', 'TID', '우울감', '고독감', '대화내용'])

    for f_num in range(file_Number):
        if f_num >= 0:
            file_name_and_time_lst = []
            for f_name in os.listdir(files_Path):
                written_time = os.path.getctime(
                    os.path.join(files_Path, f_name))
                file_name_and_time_lst.append((f_name, written_time))

            sorted_file_lst = sorted(
                file_name_and_time_lst, key=lambda x: x[1], reverse=True)
            recent_file = sorted_file_lst[f_num]
            recent_file_name = recent_file[0]

            # 엑셀 파일 로드
            data_Origin = pd.read_excel(
                os.path.join(files_Path, recent_file_name))

            # 열 이름 확인
            print("엑셀 파일 열 이름:", data_Origin.columns)

            # 'TID' 열이 있는지 확인
            if 'TID' in data_Origin.columns:
                for i in range(len(data_Origin['TID'])):
                    DATE = data_Origin['사용일'][i]
                    TID = data_Origin['TID'][i]
                    DIP = data_Origin['우울감'][i]
                    LON = data_Origin['고독감'][i]

                    if DIP >= 1 or LON >= 1:
                        data = {'사용일': DATE, 'TID': TID,
                                '우울감': DIP, '고독감': LON}
                        data_Extract = data_Extract.append(
                            data, ignore_index=True)
            else:
                print(f"Error: 'TID' 열이 파일 {recent_file_name}에 존재하지 않습니다.")
        else:
            break

    stemp = time.time()
    print(f"데이터 분류 종료까지 {stemp - start:.1f} 초")
    ################################## 데이터 분류 끝 ##############################################

    ################################## 데이터 검색 시작 ############################################

    print("데이터 검색 시작")

    # 헤더 설정
    session_Header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'}

    ####################################### 쿠키 설정 ########################################
    cookies = driver_Happy.get_cookies()
    session = requests.Session()

    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    #########################################################################################

    Total_Number = len(data_Extract['TID'])

    # 데이터 검색 시작
    print("총 인원 : {}명".format(Total_Number))

    for x in range(len(data_Extract['TID'])):
        print("현재 : {}/{}번째".format(x+1, Total_Number))

        TID = data_Extract['TID'][x]
        DATE = data_Extract['사용일'][x]

        #################################### RCP 정보 Requset ####################################
        # RCP정보 획득 URL
        url_RCP = 'https://happycommunity.happyconnect.co.kr/api/RM001001_EVT001.do'
        # RCP정보 획득 페이로드
        payload_RCP = {'search_type': 'tid', 'search_keyword': TID}

        # RCP정보 획득
        result_RCP = session.post(
            url=url_RCP, headers=session_Header, params=payload_RCP).json()
        row_RCP = result_RCP['list']
        data_RCP = row_RCP[0]['rcp_id']
        data_Name = row_RCP[0]['rcp_name']

        ##########################################################################################

        ################################### 발화어 정보 Requset ###################################
        # 발화어 정보 획득 URL
        url_Speak = 'https://happycommunity.happyconnect.co.kr/api/AN005001_EVT001.do'
        # 발화어 정보 획득 페이로드
        payload_Speak = {'rcp_id': data_RCP, 'rcp_name': data_Name,
                         'search_date_start': DATE, 'search_date_end': DATE}

        # 발화어 정보 획득
        result_Speak = session.post(
            url=url_Speak, headers=session_Header, params=payload_Speak).json()
        row_Speak = result_Speak['list']
        data_Speak = ""
        size_Low = len(row_Speak)

        for i in range(0, size_Low):
            if i+1 == size_Low:
                data_Speak = data_Speak + row_Speak[i]['ng_speech']
            else:
                data_Speak = data_Speak + row_Speak[i]['ng_speech'] + "\n"

        data_Extract['대화내용'][x] = data_Speak

        ##########################################################################################

    stemp = time.time()
    print(f"데이터 검색 종료까지 {stemp - start:.1f} 초")
    ################################## 데이터 검색 끝 ##############################################

    # 데이터를 엑셀 파일로 저장
    output_folder = r"C:\Users\USER\Desktop\24\심리상담대상자 추출자료"
    output_file_path = os.path.join(output_folder, '{}-{}-{}~{}-{}-{} 심리상담 필요대상자.xlsx'.format(
        start_Year, start_Month, start_Day, end_Year, end_Month, end_Day))
    data_Extract.to_excel(output_file_path)
    print(f"엑셀 파일이 생성되었습니다. 경로: {output_folder}")

except Exception as e:
    print(f"에러 발생: {e}")
    print("자세한 에러 메시지")
    traceback.print_exc()


end = time.time()
print(f"프로그램 총 소요시간 {end - start:.1f} 초")

driver_Happy.close()

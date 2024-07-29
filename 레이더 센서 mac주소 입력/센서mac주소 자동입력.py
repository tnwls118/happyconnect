from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import shutil
import time
import os
import os.path
import pandas as pd
from selenium.webdriver.common.alert import Alert

print("행복커뮤니티 페이지 열람")

# 크롬 인스턴스 시작
# msedgedriver.exe 파일의 경로로 수정해주세요
driver_path = r"C:\Users\82109\.vscode\work space\msedgedriver.exe"
driver = webdriver.Edge(executable_path=driver_path)

# 행복커뮤니티 페이지 열람
driver.get('https://happycommunity.happyconnect.co.kr/')
driver.maximize_window()
time.sleep(3)

# 아이디 입력 필드 찾기
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input")
input_field.send_keys("hc_csj")
time.sleep(2)

# 패스워드 필드 조회 및 입력
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
input_field.send_keys("2023tid^^")
time.sleep(2)

# 로그인 버튼 클릭
login_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
login_button.click()
time.sleep(2)

# 환경설정 버튼 클릭
button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/button")
button.click()
time.sleep(2)

# TID 버튼 클릭
button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/ul/li[8]/a")
button.click()
time.sleep(2)

# 엑셀 파일 경로
file_path = r'C:\Users\82109\.vscode\센서mapping자료.xlsx'

# 엑셀 파일 읽기
df = pd.read_excel(file_path)

# 데이터프레임 출력
print(df)
time.sleep(3)

# TID관리-상세검색 버튼 클릭
button = driver.find_element(
    By.XPATH, "/html/body/div[23]/div[2]/form/div/div[1]/ul[2]/li/select")
button.click()

# 드롭다운에서 "TID" 항목 선택
dropdown = Select(button)
dropdown.select_by_visible_text("TID")
time.sleep(2)

# 데이터프레임 반복 순회
for index, row in df.iterrows():
    # 1열 1행 데이터 조회 후 입력
    column1_data = row[0]
    input_field1 = driver.find_element(
        By.XPATH, "/html/body/div[23]/div[2]/form/div/div[1]/ul[2]/li/input[1]")
    input_field1.send_keys(column1_data)
    button = driver.find_element(
        By.XPATH, "/html/body/div[23]/div[2]/form/div/div[2]/button[1]")
    button.click()
    time.sleep(2)
    # 조회 데이터 상세보기 클릭
    button = driver.find_element(
        By.XPATH, "/html/body/div[23]/div[2]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr/td[7]/button")
    button.click()
    time.sleep(2)
    # 2열 1행 데이터 조회 후 입력
    column2_data = row[1]
    input_field2 = driver.find_element(
        By.XPATH, "/html/body/div[25]/div[2]/form/table/tbody/tr[3]/td/input")
    input_field2.send_keys(column2_data)
    time.sleep(2)
    # 3열 1행 데이터 조회 후 입력
    column3_data = row[2]
    input_field3 = driver.find_element(
        By.XPATH, "/html/body/div[25]/div[2]/form/table/tbody/tr[2]/td/input")
    input_field3.send_keys(column3_data)
    time.sleep(2)
    # TID 수정 버튼 클릭
    button = driver.find_element(
        By.XPATH, "/html/body/div[25]/div[2]/div/button[1]")
    button.click()
    time.sleep(2)
    # 알림 창 대기 및 처리
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

    time.sleep(10)

    try:
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

    time.sleep(10)
    input_field = driver.find_element(
        By.XPATH, "/html/body/div[23]/div[2]/form/div/div[1]/ul[2]/li/input[1]")
    input_field.clear()
    time.sleep(5)

    # 조회하는 데이터가 더 이상 없으면 종료
    if index >= len(df) - 1:
        break

# 웹 드라이버 종료
driver.quit()

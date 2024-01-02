from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import selenium.common.exceptions as S_exceptions

# edge 브라우저 인스턴스 생성
driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"
driver = webdriver.Chrome(driver_path)

# 웹 페이지 열기
driver.get("https://happycommunity.happyconnect.co.kr/")
# 웹 페이지 사이즈 확대
driver.maximize_window()
time.sleep(2)
# 아이디 입력 필드 찾기
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input")
input_field.send_keys("hc_csj")
time.sleep(2)

# 패스워드 입력 필드 찾기
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
input_field.send_keys("2023tid^^")
time.sleep(2)
# 로그인 버튼 클릭
login_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
login_button.click()
time.sleep(2)

# 방문 내역 및 알림처리내역 버튼 클릭
vt_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/h2/a")
vt_button.click()
voc_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/ul/li[1]/h3/a ")
voc_button.click()
time.sleep(3)

# 지역 대구분 드롭다운 클릭
dropdown = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[1]")
dropdown.click()
time.sleep(2)

# 옵션 선택
option = driver.find_element(By.XPATH, "//option[contains(text(), '지자체')]")
option.click()
time.sleep(2)

# 지역 중구분 드롭다운 클릭
dropdown = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[2]")
dropdown.click()
time.sleep(2)
# 옵션선택
option = driver.find_element(By.XPATH, "//option[contains(text(), '경상남도')]")
option.click()
time.sleep(2)


# 반복할 지역 목록
try:
    regions = ['창원시', '진주시', '사천시', '김해시', '양산시', '의령군', '함안군', '고성군', '남해군', '산청군', '거창군',
               '합천군', '통영시', '밀양시', '거제시', '창녕군', '하동군', '함양군']  # 원하는 지역을 포함하는 리스트로 변경해야 합니다.
    for region in regions:
        # 지역 소구분 드롭다운 클릭
        dropdown = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[3]")
        dropdown.click()
        time.sleep(2)

        # 옵션선택
        option = driver.find_element(
            By.XPATH, f"//option[contains(text(), '{region}')]")
        option.click()
        time.sleep(2)

        # 알림유형 선택
        a_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[1]/select")
        a_button.click()
        a_button
        option = driver.find_element(
            By.XPATH, "//option[contains(text(), '체온센서')]")
        option.click()
        time.sleep(2)

        # 상태"전체"클릭
        state_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[3]/input[1]")
        state_button.click()
        time.sleep(2)

        # 달력 이미지 클릭하기
        calendar_icon = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[3]/input[1]")
        calendar_icon.click()
        time.sleep(2)

        # 달력 월 선택
        month_dropdown = driver.find_element(
            By.XPATH, "/html/body/div[2]/div/div/select[2]")
        month_dropdown.click()

        # 10월 선택
        month_option = driver.find_element(By.XPATH, "//option[text()='10월']")
        month_option.click()

        # 특정 날짜 선택하기
        target_date = driver.find_element(
            By.XPATH, "/html/body/div[2]/table/tbody/tr[1]/td[1]/a")
        target_date.click()
        time.sleep(2)

        # 두번째 달력 이미지 클릭하기
        calendar_icon = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[3]/input[2]")
        calendar_icon.click()
        time.sleep(2)

        # 달력 월 선택
        month_dropdown = driver.find_element(
            By.XPATH, "/html/body/div[2]/div/div/select[2]")
        month_dropdown.click()
        time.sleep(2)
        # 10월 선택
        month_option = driver.find_element(By.XPATH, "//option[text()='10월']")
        month_option.click()
        time.sleep(2)
        # 특정 날짜 선택하기
        target_date = driver.find_element(
            By.XPATH, "/html/body/div[2]/table/tbody/tr[5]/td[3]/a")
        target_date.click()
        time.sleep(5)

        # 상세 검색 조회 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
        login_button.click()
        time.sleep(10)

        # 엑셀 다운로드 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/div[1]/button")
        login_button.click()
        time.sleep(2)

        # 엑셀 다운로드 사유 보고서작성용 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[31]/form/div/div[2]/div/table/tbody/tr[2]/td/input")
        login_button.click()
        time.sleep(2)

        # 엑셀다운로드 사유 확인 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[31]/form/div/div[2]/div/div/button[1]")
        login_button.click()
        time.sleep(30)

        # 다운로드된 파일을 바탕화면으로 이동
        download_folder = "C:\\Users\\82109\\Desktop"
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
except Exception as e:
    print(f"에러메시지:{e}")
driver.quit()

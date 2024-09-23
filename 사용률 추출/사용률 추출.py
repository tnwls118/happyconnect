from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import os

# 프로그램 시작 시간 기록
start_time = time.time()

# chrome 브라우저 인스턴스 생성
driver_path = r"C:\Users\82109\Desktop\시스템관련\work space\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# 웹 페이지 열기
driver.get("https://happycommunity.happyconnect.co.kr/")
print("웹 페이지 열기 완료")
# 웹 페이지 사이즈 확대
driver.maximize_window()
time.sleep(2)
print("웹 페이지 사이즈 확대 완료")

# 아이디 입력 필드 찾기
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input")
input_field.send_keys("hc_csj1")
time.sleep(2)
print("아이디 입력 완료")

# 패스워드 입력 필드 찾기
input_field = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
input_field.send_keys("2024tid^^")
time.sleep(2)
print("패스워드 입력 완료")

# 로그인 버튼 클릭
login_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
login_button.click()
time.sleep(6)
print("로그인 버튼 클릭 완료")

# 메뉴 통계정보 버튼 클릭
login_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[7]/h2/a")
login_button.click()
time.sleep(2)
print("메뉴 통계정보 버튼 클릭 완료")

# 메뉴 통계정보 ⇒ AI스피커 버튼 클릭
login_button = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[7]/ul/li[1]/ul/li[2]/a")
login_button.click()
time.sleep(2)
print("메뉴 통계정보 ⇒ AI스피커 버튼 클릭 완료")

# 지역 대구분 드롭다운 클릭
dropdown = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[1]")
dropdown.click()
time.sleep(2)
print("지역 대구분 드롭다운 클릭 완료")

# 옵션 선택
option = driver.find_element(By.XPATH, "//option[contains(text(), '지자체')]")
option.click()
time.sleep(2)
print("지자체 옵션 선택 완료")

# 지역 중구분 드롭다운 클릭
dropdown = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[2]")
dropdown.click()
time.sleep(2)
print("지역 중구분 드롭다운 클릭 완료")

# 옵션선택
option = driver.find_element(By.XPATH, "//option[contains(text(), '경상남도')]")
option.click()
time.sleep(2)
print("경상남도 옵션 선택 완료")

# 반복할 지역 목록
regions = ['창원시', '거창군', '고성군', '김해시', '남해군', '밀양시', '사천시', '산청군',
           '양산시', '의령군', '진주시', '창녕군', '거제시', '통영시', '하동군', '함안군', '함양군', '합천군']

try:
    for index, region in enumerate(regions, start=1):
        print(f"진행 중: {index}/{len(regions)} - {region}")

        # 지역 소구분 드롭다운 클릭
        dropdown = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[3]")
        dropdown.click()
        time.sleep(2)

        # 옵션선택
        option = driver.find_element(
            By.XPATH, f"//option[contains(text(), '{region}')]")
        option.click()
        time.sleep(2)

        # 달력 이미지 클릭하기
        calendar_icon = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[2]/li[2]/input[1]")
        calendar_icon.click()
        time.sleep(2)

        # 첫 번째 셀렉트 박스 클릭
        month_dropdown = driver.find_element(
            By.XPATH, "/html/body/div[2]/div/div/select[1]")
        month_dropdown.click()

        # 8월 선택
        month_option = driver.find_element(
            By.XPATH, "/html/body/div[2]/div/div/select[2]")
        month_option.click()

        month_option2 = driver.find_element(By.XPATH, "//option[text()='8월']")
        month_option2.click()

        # 특정 날짜 선택하기(1일)
        target_date = driver.find_element(
            By.XPATH, "/html/body/div[2]/table/tbody/tr[1]/td[5]/a")
        target_date.click()
        time.sleep(2)

        # 두번째 달력 이미지 클릭하기
        calendar_icon = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[2]/li[2]/input[2]")
        calendar_icon.click()
        time.sleep(2)

        # 달력 월 선택
        month_dropdown = driver.find_element(
            By.XPATH, "/html/body/div[2]/div/div/select[2]")
        month_dropdown.click()
        time.sleep(2)
        # 7월 선택
        month_option5 = driver.find_element(By.XPATH, "//option[text()='8월']")
        month_option5.click()
        time.sleep(2)
        # 특정 날짜 선택하기(31일)
        target_date = driver.find_element(
            By.XPATH, "/html/body/div[2]/table/tbody/tr[5]/td[7]/a")
        target_date.click()
        time.sleep(5)
        # 상세 검색 조회 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[2]/button[1]")
        login_button.click()
        time.sleep(100)

        # 엑셀 다운로드 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[2]/div/button")
        login_button.click()
        time.sleep(2)

        # 엑셀 다운로드 사유 보고서작성용 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[28]/form/div/div[2]/div/table/tbody/tr[2]/td/label")
        login_button.click()
        time.sleep(2)

        # 엑셀다운로드 사유 확인 버튼 클릭
        login_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div[28]/form/div/div[2]/div/div/button[1]")
        login_button.click()
        time.sleep(200)

        # 다운로드된 파일을 바탕화면으로 이동
        download_folder = "C:\\Users\\82109\\Desktop"
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        print(f"{region} 지역의 엑셀파일 다운로드 및 이동 완료")
except Exception as e:
    print(f"에러메시지: {e}")

driver.quit()

# 프로그램 종료 시간 기록 및 소요 시간 계산
end_time = time.time()
elapsed_time = end_time - start_time
print(f"프로그램 소요 시간: {elapsed_time:.2f}초")

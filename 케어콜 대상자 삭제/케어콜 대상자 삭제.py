from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,  NoSuchElementException
import time

# 전역 변수
driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"
id = "hc_csj"
pw = "2024tid^^"

# 프로그램 시작 시간 기록
start_time = time.time()
print("프로그램 시작")

# chrome driver 경로 설정
service = Service(driver_path)
hc_homepage = webdriver.Chrome(service=service)
print("chrome 경로 설정")

# 행복커뮤니티 홈페이지 조회
hc_homepage.get("https://happycommunity.happyconnect.co.kr/CO000.do")
hc_homepage.maximize_window()
time.sleep(2)
print("행커 사이트 조회 완료")

# ID, PW 입력
id_path = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input")
id_path.send_keys(id)
pw_path = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
pw_path.send_keys(pw)

login_bt = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
login_bt.click()
print("로그인 완료")

time.sleep(3)

# care_call_page 진입
care_call = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[1]/span[2]")
care_call.click()
time.sleep(5)

# 대상자 관리 메뉴 진입
care_peple = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/h2/a")
care_peple.click()
time.sleep(2)
# 지역 대구분 드롭다운 클릭
dropdown = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[1]")
dropdown.click()
time.sleep(2)
dropdown_option = hc_homepage.find_element(
    By.XPATH, "//option[contains(text(), '지자체')]")
dropdown_option.click()
time.sleep(2)
print("지자체 옵션 선택 완료")

# 중구분 드롭다운 클릭
md_dropdown = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[2]")
md_dropdown.click()

# '경상남도' 옵션 클릭
md_dropdown_option = hc_homepage.find_element(
    By.XPATH, "//option[contains(text(),'경상남도')]")
md_dropdown_option.click()
print("경상남도 옵션 선택 완료")

# 대상자 관리 필터 기간 설정 및 검색
days = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[2]/div[4]")
days.click()

search_button = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
search_button.click()
print("검색 시작")
time.sleep(4)

############################################# 반복 작업########################################################################
while True:
    try:
        filter_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/form/div/table/thead/tr/th[1]/label")
        filter_button.click()

        # 삭제 버튼 클릭
        del_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/button")
        del_button.click()
        time.sleep(2)

        try:
            alert = hc_homepage.switch_to.alert
            print(alert.text)  # 알림의 텍스트를 출력
            alert.accept()  # 확인 버튼 클릭
            time.sleep(2)
            alert2 = hc_homepage.switch_to.alert
            print(alert2.text)
            alert2.accept()
            time.sleep(3)
        except TimeoutException:
            print("알림 창이 나타나지 않았습니다.")
    except (NoSuchElementException):
        print("더 이상 삭제할 내용이 없습니다.")
        break

end_time = time.time()
runtime = end_time-start_time
print(f"프로그램 총 실행 시간:{runtime}초")

hc_homepage.quit()

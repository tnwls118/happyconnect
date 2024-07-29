from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time


print("프로그램 시작")
start_time = time.time()
# 변수 설정
hc_id = "hc_csj"
hc_pw = "2024tid^^"
homepage_Path = "https://happycommunity.happyconnect.co.kr/"
driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"

# chrome driver 경로 설정
service = Service(driver_path)
hc_homepage = webdriver.Chrome(service=service)
print("chrome 경로 설정")

# 행복커뮤니티 홈페이지 조회
hc_homepage.get(homepage_Path)
hc_homepage.maximize_window()
time.sleep(2)
print("행커 사이트 조회 완료")


# 로그인 정보 입력
ID = WebDriverWait(hc_homepage, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input"))
)
ID.send_keys(hc_id)

PW = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
PW.send_keys(hc_pw)

login_button = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
login_button.click()
time.sleep(3)
print("로그인 완료")

# 대상자 관리 페이지 조회 및 지역 설정
people_menu = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/h2/a")
people_menu.click()
people_menu1 = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/ul/li[1]/h3/a")
people_menu1.click()
time.sleep(3)

# 지역 대구분 드롭다운 클릭
dropdown1 = WebDriverWait(hc_homepage, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[1]"))
)
dropdown1.click()
time.sleep(2)

# 옵션 선택
option = WebDriverWait(hc_homepage, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//option[contains(text(), '지자체')]"))
)
option.click()
time.sleep(2)

# 지역 중구분 드롭다운 클릭
dropdown2 = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[2]")
dropdown2.click()
time.sleep(2)

# 옵션 선택
option = hc_homepage.find_element(
    By.XPATH, "//option[contains(text(), '부산광역시')]")
option.click()
time.sleep(2)

# 지역 소구분 드롭다운 클릭
dropdown3 = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[3]")
dropdown3.click()
time.sleep(2)

# 옵션 선택
option = hc_homepage.find_element(
    By.XPATH, "//option[contains(text(), '북구')]")
option.click()
time.sleep(2)

# 기간 전체 설정 및 조회
days = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[2]/div[4]")
days.click()
run_state = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[1]/input[3]")
run_state.click()
search_button = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
search_button.click()
time.sleep(3)


# 현재 페이지의 모든 항목 처리


try:
    a = 1
    while True:
        print(f"{a}번째 무한루프 중!무한로프 종료 ctr+c ")
        for n in range(1, 10):
            print(f"리스트 {n}번째 항목 작업 중")
            bogy_button = hc_homepage.find_element(
                By.XPATH, f"/html/body/div[1]/div[2]/div/div[3]/form/div/table/tbody/tr[{n}]/td[14]/button")
            bogy_button.click()
            time.sleep(3)

            # 상태 변경 버튼 클릭
            state_button = WebDriverWait(hc_homepage, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[4]/div[2]/form/div/div/table/tbody/tr[10]/td/select")))
            state_button.click()

            option = WebDriverWait(hc_homepage, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), '종료')]")))
            option.click()
            time.sleep(2)

            # 설정 버튼 클릭
            set_button = hc_homepage.find_element(
                By.XPATH, "/html/body/div[1]/div[4]/div[2]/form/div/div/div/button[1]")
            set_button.click()
            time.sleep(2)

            # 확인 Alert 처리
            WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
            alert = hc_homepage.switch_to.alert
            alert.accept()
            print("첫번째 alert 처리 완료")
            time.sleep(3)
            WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
            alert2 = hc_homepage.switch_to.alert
            alert2.accept()
            print("두번째 alert 처리 완료")
            time.sleep(3)
        a += 1  # a를 1씩 증가
except KeyboardInterrupt:
    print("프로그램이 중단되었습니다.")


hc_homepage.quit()
end_time = time.time()
run_time = end_time - start_time
print("프로그램 종료")
print(f"프로그램 실행 시간: {run_time}")

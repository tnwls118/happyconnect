from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time

print("프로그램 시작")
start_time = time.time()

# 변수 설정
hc_id = "hc_csj"
hc_pw = "2024tid^^"
homepage_Path = "https://happycommunity.happyconnect.co.kr/"
driver_path = r"C:\Users\82109\Desktop\시스템관련\work space\chromedriver.exe"

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
try:
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
except Exception as e:
    print(f"로그인 실패: {e}")
    hc_homepage.quit()
    exit()

# 방문내역 - 알림처리내역 진입
try:
    red_menu = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/h2/a"))
    )
    red_menu.click()

    red_sub_menu = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/ul/li[1]/h3/a"))
    )
    red_sub_menu.click()
    time.sleep(3)
    print("알림처리내역 진입 완료")
except Exception as e:
    print(f"알림처리내역 진입 실패: {e}")
    hc_homepage.quit()
    exit()

# 알림처리내역 필터 설정
try:
    local_state = WebDriverWait(hc_homepage, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[1]"))
    )
    local_state.click()
    local_state_drop = hc_homepage.find_element(
        By.XPATH, "//option[text()='지자체']")
    local_state_drop.click()
    time.sleep(2)

    local_state2 = WebDriverWait(hc_homepage, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[2]"))
    )
    local_state2.click()
    local_state_drop2 = hc_homepage.find_element(
        By.XPATH, "//option[text()='서울특별시']")
    local_state_drop2.click()
    time.sleep(2)

    local_state3 = WebDriverWait(hc_homepage, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[4]"))
    )
    local_state3.click()
    local_state_drop3 = hc_homepage.find_element(
        By.XPATH, "//option[text()='영등포구']")
    local_state_drop3.click()
    time.sleep(1)

    state_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[3]/input[3]")
    state_button.click()
    time.sleep(1)

    days_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[2]/div[4]")
    days_button.click()
    time.sleep(1)

    search_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
    search_button.click()
    time.sleep(3)
    print("필터 설정 완료")
except Exception as e:
    print(f"필터 설정 실패: {e}")
    hc_homepage.quit()
    exit()


# 데이터프레임 반복 순회
while True:
    try:
        # 리스트 전체 항목 체크하는 버튼 클릭
        check_all_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/div/table/thead/tr/th[1]/label")
        check_all_button.click()
        time.sleep(2)
        print("리스트 전체 클릭 완료")

        # 처리완료 버튼 클릭
        complete_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/button")
        complete_button.click()
        time.sleep(4)
        print("처리완료 버튼 클릭 완료")

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

    except (NoSuchElementException, ElementClickInterceptedException):
        # 더 이상 체크할 항목이 없거나, 버튼을 클릭할 수 없는 경우 루프 종료
        print("더 이상 체크할 항목이 없거나 클릭할 수 없습니다.")
        break


hc_homepage.quit()
end_time = time.time()
run_time = end_time - start_time
print("프로그램 종료")
print(f"프로그램 실행 시간: {run_time}")

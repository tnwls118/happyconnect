from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert
import pandas as pd
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
print("chrome 경로 설정 완료")

# 행복커뮤니티 홈페이지 조회
hc_homepage.get(homepage_Path)
hc_homepage.maximize_window()
print("행커 사이트 조회 중...")

# 로그인 정보 입력
try:
    ID = WebDriverWait(hc_homepage, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input"))
    )
    ID.send_keys(hc_id)
    print("아이디 입력 완료")

    PW = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
    PW.send_keys(hc_pw)
    print("비밀번호 입력 완료")

    login_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
    login_button.click()
    print("로그인 버튼 클릭")
    print("로그인 완료")
except Exception as e:
    print(f"로그인 실패: {e}")
    hc_homepage.quit()
    exit()

# 관리자 관리 진입
try:
    print("관리자 관리 진입 중...")
    admin_menu = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[8]/h2/a"))
    )
    admin_menu.click()

    admin_menu2 = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[8]/ul/li/h3/a"))
    )
    admin_menu2.click()
    print("관리자 관리 진입 완료")
except Exception as e:
    print(f"관리자 관리 계정 진입 실패: {e}")
    hc_homepage.quit()
    exit()

state_drop = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li/select")
state_drop.click()
state_drop_option = hc_homepage.find_element(
    By.XPATH, "//option[text()='관리자ID']")
state_drop_option.click()
# 데이터 반복 순회
try:
    print("엑셀 파일 로드 중...")
    file_path = r"C:\Users\82109\Desktop\삭제리스트.xlsx"
    df = pd.read_excel(file_path)

    # 데이터프레임 출력
    print(df.head())  # 처음 몇 줄 확인

    # 데이터프레임 반복 순회
    for index, row in df.iterrows():
        column1_data = row[0]
        print(f"데이터 입력 중: {column1_data}")

        input_field1 = WebDriverWait(hc_homepage, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li/input[1]"))
        )
        input_field1.clear()
        input_field1.send_keys(column1_data)
        time.sleep(2)
        search_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
        search_button.click()
        print("검색 버튼 클릭 완료")

        time.sleep(2)
        # 전체 선택 버튼이 있는지 확인

        all_select = WebDriverWait(hc_homepage, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/form/div/table/thead/tr/th[1]"))
        )
        all_select.click()
        print("전체 선택 버튼 클릭 완료")
        time.sleep(2)

        del_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/button[1]")
        del_button.click()
        print("삭제 버튼 클릭 완료")
        time.sleep(2)

        # 첫 번째 확인 Alert 처리
        WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
        alert = hc_homepage.switch_to.alert
        alert.accept()
        print("첫 번째 alert 처리 완료")
        time.sleep(2)
        # 두 번째 확인 Alert 처리
        WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
        alert2 = hc_homepage.switch_to.alert
        alert2.accept()
        print("두 번째 alert 처리 완료")

        if index >= len(df)-1:
            break

    print("모든 데이터 처리 완료")

except Exception as e:
    print(f"반복프레임 영역 에러 발생: {e}")
    hc_homepage.quit()
    exit()

# 프로그램 종료
hc_homepage.quit()
end_time = time.time()
run_time = end_time - start_time
print("프로그램 종료")
print(f"프로그램 실행 시간: {run_time}초")

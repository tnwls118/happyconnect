from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time

print("프로그램 시작")
start_time = time.time()

# 변수 설정
hc_id = "hc_csj1"
hc_pw = "dudn1591!"
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

# 대상자 관리 페이지 조회 및 필터 설정
people_menu = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/h2/a")
people_menu.click()
people_menu1 = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/ul/li[1]/h3/a")
people_menu1.click()
time.sleep(3)

run_state = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[1]/input[2]")
run_state.click()
print("대상자 관리 페이지 조회 및 필터 설정 완료")

# 엑셀 파일 경로 설정 및 엑셀 파일 읽기
excel_file_path = r"C:\Users\82109\Desktop\통합 문서1.xlsx"
df = pd.read_excel(excel_file_path)
print(f"데이터 프레임 출력{df}")
time.sleep(3)

# 데이터프레임 반복 순회
try:
    for index, row in df.iterrows():
        print(f"{index + 1}/{len(df)}: {row[2]} 처리 중...")

        # 3열 1행 데이터 조회 후 입력
        column3_data = row[2]
        input_field3 = WebDriverWait(hc_homepage, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[4]/li/input"))
        )
        input_field3.clear()
        input_field3.send_keys(column3_data)
        time.sleep(2)
        search_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
        search_button.click()
        time.sleep(2)

        # 검색 결과 확인
        try:
            no_results = hc_homepage.find_element(
                By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/form/div/table/tbody/tr/td").text
            if "조회 결과가 없습니다." in no_results:
                print("검색 결과가 없습니다. 다음으로 이동")
                continue
        except Exception as e:
            print(f"결과가 없습니다 에러 메시지: {e}")

        bogy_button = WebDriverWait(hc_homepage, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/form/div/table/tbody/tr[1]/td[15]/button"))
        )
        bogy_button.click()
        time.sleep(3)

        # 포켓와이파이 패스워드 항목

        pk_box = hc_homepage.find_element(
            By.XPATH, "/html/body/div[1]/div[4]/div[2]/form/div/div/table/tbody/tr[17]/td[2]/span/input")
        if pk_box == "":
            print(f"{row[2]}의 포켓와이파이 패스워드가 지워져있습니다.")
            time.sleep(2)
            pass

        else:
            pk_box.get_attribute("value")
            pk_box.clear()
            print(f"{row[2]} 포켓와이파이 패스워드 삭제 처리")
            time.sleep(2)

        # 상태 변경 버튼 클릭

            state_button = WebDriverWait(hc_homepage, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/div[1]/div[4]/div[2]/form/div/div/table/tbody/tr[10]/td/select"))
            )
        state_button.click()

        option = WebDriverWait(hc_homepage, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//option[contains(text(), '종료')]"))
        )
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

        print(f"{index + 1}/{len(df)}: {row[2]} 처리 완료")

        # 조회하는 데이터가 더 이상 없으면 종료
        if index >= len(df) - 1:
            break

except Exception as e:
    print(f"에러 메시지: {e}")

hc_homepage.quit()
end_time = time.time()
run_time = end_time - start_time
print("프로그램 종료")
print(f"프로그램 실행 시간: {run_time}")

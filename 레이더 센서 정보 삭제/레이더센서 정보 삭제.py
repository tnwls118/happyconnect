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
popup = hc_homepage.find_element(
    By.XPATH, "/html/body/div[42]/div[1]/div[2]/a/img")
popup.click()
time.sleep(2)
# 생활감지센서 설정 페이지 진입
setting_button = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/button")
setting_button.click()
tid_setting_button = hc_homepage.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/ul/li[8]/a")
tid_setting_button.click()
dropbox = hc_homepage.find_element(
    By.XPATH, "/html/body/div[23]/div[2]/form/div/div[1]/ul[2]/li/select")
dropbox.click()
option = WebDriverWait(hc_homepage, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), 'TID')]")))
option.click()
print("TID 옵션 선택 완료")

# 엑셀 파일 경로 설정 및 엑셀 파일 읽기
excel_file_path = r"C:\Users\82109\Desktop\6~7월 계약 만료.xlsx"
df = pd.read_excel(excel_file_path)
print(f"데이터 프레임 출력: {df.shape[0]}개의 행이 로드됨")
time.sleep(3)

try:
    for index, row in df.iterrows():
        print(f"{index + 1}/{len(df)}: {row[2]} 처리 중...")

        # 3열 1행 데이터 조회 후 입력
        column3_data = row[2]
        input_field3 = WebDriverWait(hc_homepage, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[23]/div[2]/form/div/div[1]/ul[2]/li/input[1]"))
        )
        input_field3.clear()
        input_field3.send_keys(column3_data)
        time.sleep(5)
        print(f"{index + 1}/{len(df)}: {column3_data} 데이터 입력 완료")

        search_button = hc_homepage.find_element(
            By.XPATH, "/html/body/div[23]/div[2]/form/div/div[2]/button[1]")
        search_button.click()
        time.sleep(5)
        print("검색 버튼 클릭 완료")

        bogy = hc_homepage.find_element(
            By.XPATH, "/html/body/div[23]/div[2]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr/td[7]/button")
        bogy.click()
        print("상세 정보 클릭 완료")

        secon_box = hc_homepage.find_element(
            By.XPATH, "/html/body/div[25]/div[2]/form/table/tbody/tr[2]/td/input")
        time.sleep(5)
        try:
            no_results = secon_box.get_attribute("value")
            if no_results == "":
                print(
                    f"{row[2]}: 2번째 박스에서 문구가 이미 삭제되어 있음. 3번째 박스로 넘어감.")
            else:
                # 두 번째 박스에 문구가 있는 경우 삭제 작업 수행
                secon_box.clear()
                time.sleep(3)
                print(f"{row[2]}: 2번째 박스 문구 삭제 완료")
        except Exception as e:
            print(f"에러 메시지: {e}")

        three_box = hc_homepage.find_element(
            By.XPATH, "/html/body/div[25]/div[2]/form/table/tbody/tr[3]/td/input")
        time.sleep(5)
        try:
            no_results1 = three_box.get_attribute("value")
            if no_results1 == "":
                print(f"{row[2]}: 3번째 박스에서 문구가 이미 삭제되어 있음. 다음 TID 조회.")
                time.sleep(2)
                time.sleep(2)
                close_button = hc_homepage.find_element(
                    By.XPATH, "/html/body/div[25]/div[2]/div/button[2]")
                close_button.click()
                continue

            else:
                # 세 번째 박스에 문구가 있는 경우 삭제 작업 수행
                three_box.clear()
                time.sleep(5)
                print(f"{row[2]}: 3번째 박스 문구 삭제 완료")

                # update 버튼 클릭
                update_button = hc_homepage.find_element(
                    By.XPATH, "/html/body/div[25]/div[2]/div/button[1]")
                update_button.click()
                time.sleep(5)
                print(f"{index + 1}/{len(df)}: 수정 버튼 클릭 완료")

                # 확인 Alert 처리
                WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
                alert = hc_homepage.switch_to.alert
                alert.accept()
                print(f"{index + 1}/{len(df)}: 첫번째 alert 처리 완료")
                time.sleep(5)
                WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
                alert2 = hc_homepage.switch_to.alert
                alert2.accept()
                print(f"{index + 1}/{len(df)}: 두번째 alert 처리 완료")
                time.sleep(5)
                print(f"{index + 1}/{len(df)}: {row[2]} 센서 정보 삭제 완료.")
        except Exception as e:
            print(f"에러 메시지: {e}")

except Exception as e:
    print(f"루프 외부에서 발생한 에러 메시지: {e}")

hc_homepage.quit()
end_time = time.time()
run_time = end_time - start_time
print("프로그램 종료")
print(f"프로그램 실행 시간: {run_time}")

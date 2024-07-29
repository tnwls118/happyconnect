from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import traceback

nugu_ID = "happytoktok02@gmail.com"
nugu_PW = "eco0401!"
homepage_Path = "https://developers.nugu.co.kr/"
chrome_driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"

try:
    print("NUGU DEV 로그인 시작")

    # Chrome 드라이버 서비스 생성
    service = Service(chrome_driver_path)

    # Chrome 드라이버 생성
    driver_Dev = webdriver.Chrome(service=service)
    driver_Dev.maximize_window()
    driver_Dev.get(homepage_Path)
    time.sleep(2)

    # NUGU Console 이동
    driver_Dev.find_element(
        By.XPATH, '/html/body/div/div[4]/header/div[2]/div/button[1]/span').click()
    time.sleep(2)

    # NUGU Dev 로그인 시작
    email_input = driver_Dev.find_element(By.ID, "userId")
    email_input.send_keys(nugu_ID)

    time.sleep(3)

    print("NUGU DEV 로그인 완료")
    time.sleep(3)
except Exception as e:
    print("NUGU DEV 로그인 오류")
    print("자세한 에러 메시지:", e)
    traceback.print_exc()
finally:
    driver_Dev.quit()  # 드라이버를 종료하여 자원을 해제합니다.

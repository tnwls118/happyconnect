from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# edge 브라우저 인스턴스 생성
chrome_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"
driver = webdriver.Chrome(chrome_path)

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
time.sleep(6)

# 케어콜 페이지 전환
care_call = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/div[1]/span[2]")
care_call.click()
time.sleep(6)

# 대상자관리 선택
member_menu = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[1]/h2/a")
member_menu.click()
time.sleep(4)

# 대상자 관리 검색 set
one_set = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[1]")
one_set.click()
one_set_option = driver.find_element(
    By.XPATH, "//option[contains(text(), '지자체')]")
one_set_option.click()
time.sleep(1)
two_set = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[1]/li/select[2]")
two_set.click()
two_set_option = driver.find_element(
    By.XPATH, "//option[contains(text(), '경상남도')]")
two_set_option.click()
time.sleep(1)

# 대상자 관리 검색_이용상태 설정
member_state = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[1]/input[2]")
member_state.click()

# 대상자 관리 검색 클릭
sc_bt = driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[2]/button[1]")
sc_bt.click()

time.sleep(6)

# 대상자 삭제 반복
start_time = time.time()
while True:
    try:
        select_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/form/div/table/thead/tr/th[1]/label")
        select_button.click()

        del_button = driver.find_element(
            By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/button")
        del_button.click()

        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = Alert(driver)
        alert.accept()

        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = Alert(driver)
        alert.accept()

        time.sleep(3)

    except NoSuchElementException:
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"작업 완료(총 소요시간: {elapsed_time}초)")
        break

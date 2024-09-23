from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

print("프로그램 시작")
start_time = time.time()

# 변수 설정
hc_id = "hc_csj1"
hc_pw = "dudn1591!"
homepage_Path = "https://happycommunity.happyconnect.co.kr/"
driver_path = r"C:\Users\82109\Desktop\시스템관련\work space\chromedriver.exe"

# Chrome driver 경로 설정
service = Service(driver_path)
hc_homepage = webdriver.Chrome(service=service)
print("Chrome 경로 설정 완료")

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
    time.sleep(2)

    PW = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input")
    PW.send_keys(hc_pw)
    print("비밀번호 입력 완료")
    time.sleep(2)

    login_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button")
    login_button.click()
    print("로그인 버튼 클릭 완료")
    time.sleep(2)
except Exception as e:
    print(f"로그인 실패: {e}")
    hc_homepage.quit()
    exit()

# 알림창 삭제
try:
    dele = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[42]/div[1]/div[2]/a/img"))
    )
    dele.click()
    print("알림창 닫기 완료")
except TimeoutException:
    print("알림창을 찾을 수 없습니다.")

# 방문내역 진입 및 필터 설정
try:
    visit_menu = WebDriverWait(hc_homepage, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/h2/a"))
    )
    visit_menu.click()
    print("방문 메뉴 클릭 완료")
    time.sleep(4)

    visit_menu2 = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[4]/ul/li[1]/h3/a")
    visit_menu2.click()
    print("방문 메뉴2 클릭 완료")
    time.sleep(3)

    state = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[2]/li[3]/input[3]")
    state.click()
    print("상태 선택 완료")
    time.sleep(3)

    calendar_icon1 = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[3]/input[1]")
    calendar_icon1.click()
    print("첫번째 달력아이콘 클릭 완료")
    time.sleep(2)

    year_dropdown1 = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/div/div/select[1]")
    year_dropdown1.click()
    print("연도 드롭다운 클릭 완료")

    year_option1 = hc_homepage.find_element(
        By.XPATH, "//option[text()='2022']")
    year_option1.click()
    print("2022년 선택 완료")

    time.sleep(2)
    month_dropdown1 = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/div/div/select[2]")
    month_dropdown1.click()
    print("월 드롭다운 클릭 완료")

    time.sleep(2)
    month_option1 = hc_homepage.find_element(
        By.XPATH, "//option[text()='1월']")
    month_option1.click()
    print("1월 선택 완료")

    time.sleep(2)
    target_date1 = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/table/tbody/tr[1]/td[7]/a")
    target_date1.click()
    print("1일 선택 완료")

    time.sleep(2)

    calendar_icon = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/form/div/div[1]/ul[3]/li[3]/input[2]")
    calendar_icon.click()
    print("달력 아이콘 클릭 완료")

    time.sleep(2)
    year_dropdown = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/div/div/select[1]")
    year_dropdown.click()
    print("연도 드롭다운 클릭 완료")

    year_option = hc_homepage.find_element(By.XPATH, "//option[text()='2022']")
    year_option.click()
    print("2022년 선택 완료")

    time.sleep(2)
    month_dropdown = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/div/div/select[2]")
    month_dropdown.click()
    print("월 드롭다운 클릭 완료")

    time.sleep(2)
    month_option5 = hc_homepage.find_element(
        By.XPATH, "//option[text()='12월']")
    month_option5.click()
    print("12월 선택 완료")

    time.sleep(2)
    target_date = hc_homepage.find_element(
        By.XPATH, "/html/body/div[2]/table/tbody/tr[5]/td[7]/a")
    target_date.click()
    print("31일 선택 완료")

    time.sleep(2)
    login_button = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[2]/button[1]")
    login_button.click()
    print("상세 검색 조회 버튼 클릭 완료")
    time.sleep(10)
except Exception as e:
    print(f"알림처리내역 필터 설정 에러: {e}")

# 반복작업영역
try:
    list_number = hc_homepage.find_element(
        By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/div[2]/div/select")
    list_number.click()
    print("목록 개수 선택 완료")

    list_option = hc_homepage.find_element(
        By.XPATH, "//option[text()='30개씩 보기']")
    list_option.click()
    time.sleep(10)
    print("30개씩 보기 옵션 선택 완료")

    current_page = 1

    while True:
        try:
            # 로더가 보이지 않을 때까지 대기
            WebDriverWait(hc_homepage, 200).until(
                EC.invisibility_of_element_located(
                    (By.ID, 'loaderWrapperDataTable'))
            )

            element_to_click = WebDriverWait(hc_homepage, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/div[1]/div[2]/div/div[3]/div/table/thead/tr/th[1]/label'))
            )

            hc_homepage.execute_script(
                "arguments[0].click();", element_to_click)
            print("첫 번째 요소 클릭 완료")
            time.sleep(2)

            button_to_click = WebDriverWait(hc_homepage, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/div[1]/div[2]/div/div[4]/button'))
            )
            hc_homepage.execute_script(
                "arguments[0].click();", button_to_click)
            print("두 번째 버튼 클릭 완료")
            time.sleep(2)

            WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
            alert = hc_homepage.switch_to.alert
            alert.accept()
            print("첫 번째 alert 처리 완료")
            time.sleep(2)

            WebDriverWait(hc_homepage, 10).until(EC.alert_is_present())
            alert2 = hc_homepage.switch_to.alert
            alert2.accept()
            print("두 번째 alert 처리 완료")
            time.sleep(2)

        except TimeoutException:
            print("첫 번째 요소를 클릭할 수 없거나 요소가 없습니다. 다음 페이지로 이동합니다.")
            current_page += 1

            if current_page % 10 == 0:
                try:
                    next_page_button = WebDriverWait(hc_homepage, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "/html/body/div[1]/div[2]/div/div[3]/div/div[2]/a[3]"))
                    )
                    next_page_button.click()
                    print(f"{current_page} 페이지 이후로 이동 완료")
                    time.sleep(2)
                except TimeoutException:
                    print("추가 페이지 이동 버튼이 없습니다. 작업을 종료합니다.")
                    break
            else:
                next_page_xpath = f'/html/body/div[1]/div[2]/div/div[3]/div/div[2]/span/a[{current_page % 10 + 1}]'

                try:
                    next_page_button = WebDriverWait(hc_homepage, 10).until(
                        EC.element_to_be_clickable((By.XPATH, next_page_xpath))
                    )
                    next_page_button.click()
                    print(f"{current_page} 페이지로 이동 완료")
                    time.sleep(2)
                except TimeoutException:
                    print("더 이상 이동할 페이지가 없습니다. 작업을 종료합니다.")
                    break

except Exception as e:
    print(f"반복작업 중 에러 발생: {e}")

# 프로그램 종료
hc_homepage.quit()
end_time = time.time()
print(f"프로그램 종료. 소요 시간: {end_time - start_time} 초")

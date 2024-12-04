from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
input_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[1]/dd/input"))
)
input_field.send_keys("hc_csj1")
print("아이디 입력 완료")

# 패스워드 입력 필드 찾기
input_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/dl[2]/dd/input"))
)
input_field.send_keys("dudn1591!")
print("패스워드 입력 완료")

# 로그인 버튼 클릭
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div/div[2]/form/div[1]/button"))
)
login_button.click()
print("로그인 버튼 클릭 완료")

# 메뉴 통계정보 버튼 클릭
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[7]/h2/a"))
)
login_button.click()
print("메뉴 통계정보 버튼 클릭 완료")

# 메뉴 통계정보 ⇒ AI스피커 버튼 클릭
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/div/ul/li[7]/ul/li[1]/ul/li[2]/a"))
)
login_button.click()
print("메뉴 통계정보 ⇒ AI스피커 버튼 클릭 완료")

# 지역 대구분 드롭다운 클릭
dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[1]"))
)
dropdown.click()
print("지역 대구분 드롭다운 클릭 완료")

# 옵션 선택
option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), '지자체')]"))
)
option.click()
print("지자체 옵션 선택 완료")

# 지역 중구분 드롭다운 클릭
dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[2]"))
)
dropdown.click()
print("지역 중구분 드롭다운 클릭 완료")

# 옵션 선택
option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//option[contains(text(), '경상남도')]"))
)
option.click()
print("경상남도 옵션 선택 완료")

# # 옵션 선택
# option = WebDriverWait(driver, 10).until(
#     EC.element_to_be_clickable(
#         (By.XPATH, "//option[contains(text(), '부산광역시')]"))
# )
# option.click()
# print("부산광역시 옵션 선택 완료")

# 반복할 지역 목록
regions = ['거제시', '거창군', '고성군', '김해시', '남해군', '밀양시', '사천시', '산청군',
           '양산시', '의령군', '진주시', '창녕군', '창원시', '통영시', '하동군', '함안군', '함양군', '합천군']

# # 반복할 지역 목록
# regions = ['중구', '부산진구', '사하구', '동구', '북구', '강서구']

try:
    for index, region in enumerate(regions, start=1):
        print(f"진행 중: {index}/{len(regions)} - {region}")

        # 지역 소구분 드롭다운 클릭
        dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[1]/li/select[3]"))
        )
        dropdown.click()

        # 옵션 선택
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//option[contains(text(), '{region}')]"))
        )
        option.click()
        print(f"{region} 지역 선택 완료")

        # 달력 이미지 클릭하기
        calendar_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[2]/li[2]/input[1]"))
        )
        calendar_icon.click()

        # 첫 번째 셀렉트 박스 클릭
        month_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/div/div/select[1]"))
        )
        month_dropdown.click()

        # 11월 선택
        month_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/div/div/select[2]"))
        )
        month_option.click()

        month_option2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//option[text()='11월']"))
        )
        month_option2.click()

        # 특정 날짜 선택하기(1일)
        target_date = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/table/tbody/tr[1]/td[6]/a"))
        )
        target_date.click()
        print("시작 날짜 선택 완료")

        # 두 번째 달력 이미지 클릭하기
        calendar_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[1]/ul[2]/li[2]/input[2]"))
        )
        calendar_icon.click()

        # 달력 월 선택
        month_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/div/div/select[2]"))
        )
        month_dropdown.click()

        # 11월 선택
        month_option5 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//option[text()='11월']"))
        )
        month_option5.click()

        # 특정 날짜 선택하기(30일)
        target_date = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[2]/table/tbody/tr[5]/td[7]/a"))
        )
        target_date.click()
        print("종료 날짜 선택 완료")

        time.sleep(3)

        # 상세 검색 조회 버튼 클릭
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[1]/div[2]/button[1]"))
        )
        search_button.click()

        time.sleep(60)

        # 엑셀 다운로드 버튼 클릭
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[1]/form/div[2]/div/button"))

        )
        download_button.click()
        print("엑셀 다운로드 버튼 클릭 완료")

        time.sleep(3)

        # 엑셀 다운로드 사유 보고서작성용 버튼 클릭
        reason_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[33]/form/div/div[2]/div/table/tbody/tr[2]/td/input"))
        )
        reason_button.click()

        time.sleep(2)

        # 엑셀다운로드 사유 확인 버튼 클릭
        confirm_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[2]/div[33]/form/div/div[2]/div/div/button[1]"))
        )
        confirm_button.click()

        time.sleep(120)

        # 다운로드된 파일을 바탕화면으로 이동
        download_folder = "C:\\Users\\82109\\Desktop"
        print(f"{region} 지역의 엑셀파일 다운로드 및 이동 완료")
except Exception as e:
    print(f"에러메시지: {e}")

driver.quit()

# 프로그램 종료 시간 기록 및 소요 시간 계산
end_time = time.time()
elapsed_time = end_time - start_time
print(f"프로그램 소요 시간: {elapsed_time:.2f}초")

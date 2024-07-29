import time
import traceback
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains


# 주요 변수 모음
chrome_driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"
homepage_path = r"https://memorywalk.kr/login"
id = "admin"
pw = "skt@1234kk"

# 프로그램 실행 시간
start_time = time.time()

print("프로그램 시작", start_time)

# Chrome 드라이버 서비스 생성
service = Service(chrome_driver_path)

# 기억산책 홈페이지 조회
driver = webdriver.Chrome(service=service)
driver.maximize_window()
driver.get(homepage_path)
time.sleep(2)

try:
    print("로그인 시작")
    # 기억산책 ID 및 패스워드 입력
    id_box = driver.find_element(
        By.XPATH, "/html/body/div/div/div[1]/form/div[2]/input")
    id_box.send_keys(id)
    pw_box = driver.find_element(
        By.XPATH, "/html/body/div/div/div[1]/form/div[3]/input")
    pw_box.send_keys(pw)
    # 옵션 선택
    option = driver.find_element(
        By.XPATH, "/html/body/div/div/div[1]/form/div[4]/div/div[1]/input")
    option.click()
    option_1 = driver.find_element(By.CSS_SELECTOR, "#selectOrg-opt-4")
    option_1.click()
    login_bt = driver.find_element(
        By.XPATH, "/html/body/div/div/div[1]/form/div[6]/div[2]/button")
    login_bt.click()
    print("로그인 성공")
except:
    print("로그인 실패")
    print("에러 메시지:", traceback.format_exc())

time.sleep(5)

# 총 페이지 수
total_pages = 58

# # 수집된 데이터를 저장할 리스트 초기화
# collected_data = []

# 페이지를 순차적으로 탐색
print("데이터 순차적 탐색:")
for page in range(1, total_pages + 1):
    # 각 페이지에서 아이템의 개수 가져오기
    items_on_page = len(driver.find_elements(
        By.XPATH, '//*[@id="dataTable"]/tbody/tr'))
    for item_index in range(1, items_on_page + 1):  # 아이템의 개수만큼 반복
        try:
            # 성별 정렬
            list_set = driver.find_element(
                By.XPATH, "/html/body/div[1]/div[2]/main/div[1]/div/div[2]/div/div/div/div/div[2]/div/table/thead/tr[1]/th[4]")
            list_set.click()
            time.sleep(2)
            # 아이디를 클릭 (XPath는 실제 페이지 구조에 맞게 수정 필요)
            item_xpath = f'//*[@id="dataTable"]/tbody/tr[{item_index}]/td[3]'
            item = driver.find_element(By.XPATH, item_xpath)
            item.click()
            time.sleep(2)

            print(f"{items_on_page}개 중 {item_index}번째 아이템을 클릭했습니다.")

            # # 기간통계 클릭(기간설정 '24년 1월 ~ 6월)
            # bt1 = driver.find_element(
            #     By.XPATH, "/html/body/div[1]/div[2]/main/div[1]/div[2]/div/a[2]")
            # bt1.click()
            # time.sleep(5)

            # # 첫 번째 드롭다운 선택
            # calendar_icon = driver.find_element(
            #     By.XPATH, "/html/body/div[1]/div[2]/main/div[1]/div[3]/div[2]/div[1]/div/input[1]")
            # calendar_icon.click()
            # # 달력에서 "1" 선택
            # month_option = driver.find_element(
            #     By.XPATH, "//option[text()='1']")
            # month_option.click()

            # # 두 번째 드롭다운 선택
            # calendar_icon1 = driver.find_element(
            #     By.XPATH, "/html/body/div[1]/div[2]/main/div[1]/div[3]/div[2]/div[1]/div/input[2]")
            # calendar_icon1.click()
            # # 달력에서 "1" 선택
            # month_option1 = driver.find_element(
            #     By.XPATH, "//option[text()='6']")
            # month_option1.click()
            # # 실제로 필요한 데이터의 XPath를 찾아서 수정해야 합니다.
            # data_1 = item_index
            # data_2 = driver.find_element(By.XPATH, '//*[@id="answer"]').text
            # data_3 = driver.find_element(
            #     By.XPATH, '//*[@id="solvingTime"]').text
            # data_4 = driver.find_element(
            #     By.XPATH, '//*[@id="correctRate"]').text

            # collected_data.append([data_1, data_2, data_3, data_4])
            # 뒤로가기
            driver.back()
            time.sleep(2)

        except Exception as e:
            print(f"오류 발생: {e}")
            continue

    # 다음 페이지로 이동 (XPath는 실제 페이지 구조에 맞게 수정 필요)
    try:
        next_page_button = driver.find_element(
            By.XPATH, '/html/body/div[1]/div[2]/main/div[1]/div/div[2]/div/div/div/div/div[3]/div[2]/div/ul/li[9]/a')
        next_page_button.click()
        time.sleep(2)
        print("다음페이지로 이동했습니다.")
    except Exception as e:
        print(f"페이지 이동 오류: {e}")
        break

# # 수집한 데이터를 데이터프레임으로 변환
# df = pd.DataFrame(collected_data, columns=[
#                   'ID', '기간내 문제 풀이 수', '기간내 문제 풀이 시간', '기간내 평균 정답률'])

# # 엑셀 파일로 저장
# # 사용자의 다운로드 폴더 경로 가져오기
# download_dir = str(Path.home() / "Downloads")
# # 엑셀 파일 경로 설정
# excel_path = os.path.join(download_dir, "기억산책_데이터.xlsx")
# # 엑셀 파일 저장
# df.to_excel(excel_path, index=True)
end_time = time.time()
total_time = end_time - start_time
# print(f"엑셀 파일이 저장되었습니다: {excel_path}")
print(f"프로그램 실행 시간: {total_time:.2f} 초")

# 드라이버 종료
driver.quit()

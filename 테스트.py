from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import requests
import time

# ChromeDriver 경로 설정
chrome_driver_path = r"C:\Users\82109\.vscode\work space\chromedriver.exe"  # 원시 문자열 사용

# Service 객체 생성
service = Service(chrome_driver_path)

# Selenium 웹 드라이버 설정
driver = webdriver.Chrome(service=service)
driver.get("https://happycommunity.happyconnect.co.kr/")

# 페이지 로딩 대기
time.sleep(5)  # 필요에 따라 대기 시간 조정

# 쿠키 가져오기
cookies = driver.get_cookies()

# 세션 생성
session = requests.Session()

# 쿠키를 세션에 추가
for cookie in cookies:
    session.cookies.set(cookie['name'], cookie['value'])

# RCP 정보 요청을 위한 URL
url_RCP = 'https://happycommunity.happyconnect.co.kr/api/RM001001_EVT001.do'

# RCP 정보 요청을 위한 페이로드 (테스트용 TID 값)
TID = 'kh225@hs.com'  # 테스트할 TID 값
payload_RCP = {'search_type': 'tid', 'search_keyword': TID}

# RCP 정보 요청
result_RCP = session.post(url=url_RCP, headers={
                          'User-Agent': 'Mozilla/5.0'}, params=payload_RCP)
result_RCP_json = result_RCP.json()

# RCP 정보 출력
print("RCP 정보:")
print(result_RCP_json)

# 드라이버 종료
driver.quit()

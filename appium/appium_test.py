from appium import webdriver
from time import sleep

# Desired Capabilities 설정
desired_caps = {
    "platformName": "Android",
    "platformVersion": "12.0",
    "deviceName": "Pixel_3a_API_31",
    "automationName": "UiAutomator2",
    "noReset": True
}

try:
    # Appium 서버와 연결
    driver = webdriver.Remote("http://localhost:4723", desired_caps)
    print("Successfully connected to Appium server and device")

    # 기본 화면 정보 가져오기
    sleep(5)
    screen_size = driver.get_window_size()
    print(f"Screen size: {screen_size}")

    # 화면 캡처
    screenshot_path = "screenshot.png"
    driver.save_screenshot(screenshot_path)
    print(f"Screenshot saved to {screenshot_path}")

except Exception as e:
    print(f"Error: {e}")

finally:
    # 드라이버 종료
    if 'driver' in locals():
        driver.quit()

from select import select
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from imaplib import IMAP4
from bs4 import BeautifulSoup as BS
from google.cloud import vision
import cv2
import numpy as np
import pyzmail
import imapclient
import time
import pandas as pd
import os
import io
# from PIL import ImageGrab


# NUGU Dev 로그인/Console 이동
def login_NuguDev(driver_Dev):
    try:
        print("NUGU DEV 로그인 시작")
        nugu_ID = "happytoktok02@gmail.com"
        nugu_PW = "eco0401!"

        homepage_Path = "https://developers.nugu.co.kr/"

# ChromeDriver 실행 파일을 공식 웹 사이트에서 다운로드하고 해당 경로를 제공합니다.
        driver_path = r"C:\Users\82109\.vscode\work space\msedgedriver.exe"

# chrome driver 경로 설정
        service = Service(driver_path)
        driver_Dev = webdriver.Chrome(service=service)
        print("chrome 경로 설정")
        wait = WebDriverWait(driver_Dev, 60)

# NUGU Dev 로그인 이동
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/header/div[2]/div/div[2]/ul/li[1]/a'))).click()

# NUGU Dev 로그인 시작
        wait.until(EC.element_to_be_clickable(
            (By.ID, "userId"))).send_keys(nugu_ID)
        driver_Dev.find_element(By.ID, "password").send_keys(nugu_PW)
        driver_Dev.find_element(By.ID, "authLogin").click()
        driver_Dev.implicitly_wait(60)

# NUGU Console 이동
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/header/div[2]/div/button[1]')))
        driver_Dev.get('https://developers.nugu.co.kr/#/dashboard')
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/header/div[2]/nav/ul/li[3]/a')))

# 반환
        print("NUGU DEV 로그인 완료")
        return driver_Dev, True
    except:
        print("NUGU DEV 로그인 오류")
        return driver_Dev, False


time.sleep(5)

# NUGU Dev 대상자 삭제


def deleteUser_Nugu(driver_Dev, TID):
    try:
        print("대상자 삭제 시작")
        # 검색불가 TID영역 삭제
        cur_TID = ''
        if 'hs.com' in TID:
            cur_TID = TID.replace('hs.com', '')
        elif 'group.tid.com' in TID:
            cur_TID = TID.replace('group.tid.com', '')

        # NUGU Dev 삭제 메뉴 이동
        wait = WebDriverWait(driver_Dev, 60)
        driver_Dev.find_element(
            By.XPATH, '//*[@id="app"]/div[5]/header/div[2]/nav/ul/li[3]/a').click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="side-menu"]/div/ul/li[4]/ul/li[3]/a'))).click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="side-menu"]/div/ul/li[4]/ul/li[3]/ul/li[3]/a'))).click()

        # 삭제 대상자 조회
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[1]/div/dd/span/p/input'))).clear()
        driver_Dev.find_element(
            By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[1]/div/dd/span/p/input').send_keys(cur_TID)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[1]/div/dd/span/p/label'))).click()
        time.sleep(2)

        # 대상자 삭제(기존 등록된 대상자), 불필요시 pass
        search_Val = driver_Dev.find_element(
            By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[2]/div[1]/span/strong').text
        if search_Val == '1':
            driver_Dev.find_element(
                By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[3]/form/table/thead/tr/th[1]/div/label/span').click()
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[5]/div/article/div/div[2]/div[2]/div/button[3]'))).click()
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div/div[5]/div[1]/section/div[4]/div[2]/footer/button[2]'))).click()
        else:
            pass
        print("대상자 삭제 종료")
        return driver_Dev, True
    except:
        print("대상자 삭제 오류")
        return driver_Dev, False
# NUGU Dev 대상자 초대


def inviteUser_Nugu(driver_Dev, TID, NAME, EMAIL, PHONE):
    try:
        print("대상자 초대 시작")

        # NUGU Dev 초대 메뉴 이동
        wait = WebDriverWait(driver_Dev, 60)
        driver_Dev.find_element(
            By.XPATH, '//*[@id="app"]/div[5]/header/div[2]/nav/ul/li[3]/a').click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="side-menu"]/div/ul/li[4]/ul/li[3]/a'))).click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="side-menu"]/div/ul/li[4]/ul/li[3]/ul/li[2]/a'))).click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[2]/div/article/div[1]/button'))).click()
        time.sleep(1)

        # 그룹 선택
        driver_Dev.find_element(
            By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[2]/div/article[2]/div/div[1]/div/table/tbody/tr/td/div/label[3]/span').click()
        # wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[2]/div/article[2]/div/div[1]/div/table/tbody/tr/td/div/label[3]/span"]'))).click()

        # 초대하기 - Play선택하기 클릭
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[2]/div/article[3]/div/div[2]/button'))).click()
        time.sleep(1)

        # Play 선택(기억검사, 좋은생각, 마음체조, 생활감지센서, 두뇌톡톡, 소식톡톡)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[1]/td[1]/div/label/span'))).click()
        # driver_Dev.find_element(By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[1]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[2]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[3]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[4]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[5]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[8]/td[1]/div/label/span""").click()
        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/div/div/div/form/table/tbody/tr[9]/td[1]/div/label/span""").click()
        time.sleep(1)

        driver_Dev.find_element(
            By.XPATH, """//*[@id="bizInvitationPlaySelect"]/div[2]/footer/button[2]""").click()
        time.sleep(1)

        # 초대 TID 등록
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="name0"]'))).send_keys(NAME)
        # driver_Dev.find_element(By.XPATH, """//*[@id="name0"]""").send_keys(NAME)
        driver_Dev.find_element(
            By.XPATH, """//*[@id="email0"]""").send_keys(EMAIL)
        driver_Dev.find_element(
            By.XPATH, """//*[@id="phone0"]""").send_keys(PHONE)
        driver_Dev.find_element(By.XPATH, """//*[@id="tag0"]""").send_keys(TID)
        time.sleep(1)

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[5]/div[1]/section/div[3]/button[2]'))).click()
        # driver_Dev.find_element(By.XPATH, """//*[@id="app"]/div[5]/div[1]/section/div[3]/button[2]""").click()
        # time.sleep(1)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div/div[5]/div[1]/div[1]/div[2]/footer/button'))).click()
        # driver_Dev.find_element(By.XPATH, """/html/body/div/div[5]/div[1]/div[1]/div[2]/footer/button""").click()
        # time.sleep(1)

        print("대상자 초대 성공")
        return driver_Dev, True
    except:
        print("대상자초대 오류")
        return driver_Dev, False
# GMAIL 로그인


def login_Gmail(ID, PW):
    try:
        # 로그인
        imap_obj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
        imap_obj.login(ID, PW)
        print("GMAIL 로그인 완료")
        return imap_obj, True
    except:
        print("GMAIL로그인 오류")
        return imap_obj, False
# GMAIL 비우기


def empty_Gmail(imap_obj):
    try:
        print("GMAIL 비우기 시작")
        time.sleep(1)
        # 받은 편지함 이동
        imap_obj.select_folder('INBOX', readonly=False)
        UIDs = imap_obj.search(['ALL'])
        imap_obj.delete_messages(UIDs)
        imap_obj.expunge()
        print("GMAIL 비우기 완료")
        return imap_obj, True
    except:
        print("GMAIL비우기 오류")
        return imap_obj, False
# GMAIL NUGU 초대 링크 확인


def inviteLink_Gmail(imap_obj):
    print("GMAIL NUGU초대 링크 확인 시작")
    invite_URL = ''
    time.sleep(5)

    # 받은 편지함 이동
    imap_obj.select_folder('INBOX', readonly=True)

    # NUGU메일 검색
    UID = imap_obj.search(['FROM', 'no-reply@nugu.co.kr'])

    # 메일 내용 검색하여 초대 URL출력
    raw_msg = imap_obj.fetch(UID, ['BODY[]'])
    for i in raw_msg:
        msg = pyzmail.PyzMessage.factory(raw_msg[i][b'BODY[]'])
        text = msg.html_part.get_payload().decode(msg.html_part.charset)
        soup = BS(text, "html.parser")
        invite_URL = soup.find("a")["href"]

    print("GMAIL NUGU초대 링크 확인 완료")
    return imap_obj, invite_URL
# 외부 초대 링크 수락


def AcceptInvite_Nugu(URL, TID, PW):
    try:
        print("누구 외부 초대 실행")
        # 초대 수락 창 실행
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # driver_Invite = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver_Invite = webdriver.Chrome(options=options)
        driver_Invite.maximize_window()
        driver_Invite.get(URL)
        wait = WebDriverWait(driver_Invite, 60)

        # 초대 수락 이동
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[2]/div/section/div/div[1]/div[3]/button')))
        driver_Invite.find_element(
            By.XPATH, '/html/body/div[2]/div/section/div/div[1]/div[3]/button').click()
        wait = WebDriverWait(driver_Invite, 60)

        # NUGU Invite 로그인 시작
        wait.until(EC.element_to_be_clickable((By.ID, 'userId')))
        driver_Invite.find_element(By.ID, "userId").send_keys(TID)
        driver_Invite.find_element(By.ID, "password").send_keys(PW)
        driver_Invite.find_element(By.ID, "authLogin").click()
        driver_Invite.implicitly_wait(60)

        time.sleep(5)
        # ★★★★★[중요]이미지 처리★★★★★
        try:
            while 1:
                cur_URL = driver_Invite.current_url
                try:
                    if cur_URL.find("authorize.do?") != -1:
                        print("문자열 검증 시작")

                        # 기존 이미지 삭제 Section
                        if ImgCheck("./comp_Img.png") == True:
                            os.remove("./comp_Img.png")
                            print("기존 다운로드 이미지 삭제 완료")
                        if ImgCheck("./cvt_Img.png") == True:
                            os.remove("./cvt_Img.png")
                            print("기존 흑백 이미지 삭제 완료")

                        time.sleep(1)
                        wait.until(EC.element_to_be_clickable(
                            (By.XPATH, '//*[@id="reload"]/span'))).click()
                        time.sleep(1)

        #                 # 이미지 다운로드
        #                 driver_Invite, download_Result = CapImgDown(
        #                     driver_Invite)

        #                 if download_Result == True:
        #                     # 이미지 흑백반전
        #                     try:
        #                         cvt_Result = Change_ImgB2H()
        #                         if cvt_Result == True:
        #                             # 이미지 흑백 변환 완료시 구글 이미지 Read시작
        #                             try:
        #                                 text_Result = ReadImageFromGoogle()
        #                                 if text_Result != "":
        #                                     # 입력부
        #                                     time.sleep(2)
        #                                     driver_Invite.find_element(
        #                                         By.XPATH, '//*[@id="answer"]').clear()
        #                                     driver_Invite.find_element(
        #                                         By.XPATH, '//*[@id="answer"]').send_keys(text_Result)
        #                                     time.sleep(1)
        #                                     driver_Invite.find_element(
        #                                         By.ID, "authLogin").click()
        #                                     time.sleep(3)
        #                                 else:
        #                                     pass
        #                             except:
        #                                 pass
        #                         else:
        #                             pass
        #                     except:
        #                         pass
        #                 else:
        #                     pass
        #             else:
        #                 break
                except:
                    pass
        except:
            pass
        # 3개월 비밀번호 갱신
        try:
            cur_URL = driver_Invite.current_url
            if cur_URL.find("passwordstate") != -1:
                driver_Invite.find_element(By.ID, "nextChange").click()
                driver_Invite.implicitly_wait(60)
        except:
            pass
        # 서비스동의
        try:
            cur_URL = driver_Invite.current_url
            if cur_URL.find("member") != -1:
                driver_Invite.find_element(
                    By.XPATH, """//*[@id="termsAgreeform"]/div[1]/div[1]/span/label""").click()
                time.sleep(0.5)
                driver_Invite.find_element(
                    By.XPATH, """//*[@id="btn_next"]""").click()
                driver_Invite.implicitly_wait(60)
        except:
            pass

        # 창 종료
        driver_Invite.close()
        print("누구 외부 초대 성공")
        return True
    except:
        print("누구 외부초대 오류")
        return False
# 이미지 다운로드 실행


# def CapImgDown(Driver):
#     try:
#         img = ImageGrab.grab((723, 757, 1114, 869))
#         # img = ImageGrab.grab((723, 688, 1115, 797))
#         img.save("comp_Img.png")
#         print("이미지 다운로드 완료")
#         return Driver, True
#     except:
#         print("이미지 다운로드 실패")
#         return Driver, False

#     return Driver
# 이미지 글자 Google Could 읽기


def ReadImageFromGoogle():
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'checkchar-5467c0fb1716.json'

    client_options = {'api_endpoint': 'eu-vision.googleapis.com'}
    client = vision.ImageAnnotatorClient(client_options=client_options)

    path = './cvt_Img.png'
    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.text_detection(image=image)
    texts = response.text_annotations
    result_Return = ''
    for text in texts:
        print(text.description)
        result_Return = text.description
        break
    print("문자 Google판독 완료")
    return result_Return
# Comp_Img Check


def ImgCheck(STR):
    if os.path.isfile(STR) == True:
        return True
    else:
        return False
# 이미지 흑백 반전


def Change_ImgB2H():
    if ImgCheck('./comp_Img.png') == True:

        img = cv2.imread('comp_Img.png')  # 파일명

        canvas = np.zeros(shape=img.shape, dtype=np.uint8)
        canvas.fill(255)
        canvas[np.where((img == [0, 0, 0]).all(axis=2))] = [0, 0, 0]
        canvas[np.where((img == [64, 64, 64]).all(axis=2))] = [0, 0, 0]
        canvas[np.where((img == [255, 255, 255]).all(axis=2))] = [0, 0, 0]
        canvas[np.where((img == [192, 192, 192]).all(axis=2))] = [0, 0, 0]

        cv2.imwrite('cvt_Img.png', canvas)
        print("이미지 흑백 반전 완료")
        return True
    else:
        print("흑백반전 실패")
        return False


if __name__ == "__main__":
    # iMAP크기제한 해제
    IMAP4._MAXLINE = 10000000
    # GMAIL ID/PW(앱 비밀번호) 설정
    gmail_ID = 'happytoktok02@gmail.com'
    gmail_PW = 'ekhpwlnrlouxxatf'
    imap_Gmail = ''
    invite_URL = ''
    driver_Main = ""
    result_NoMail = False
    count_Nomail = 0

    try:
        print("프로그램 시작지점")
        # NUGU DEV 로그인
        while 1:
            driver_Main, result_Main = login_NuguDev(driver_Main)
            if result_Main == True:
                break

        # Gmail로그인
        while 1:
            imap_Gmail, result_LogGM = login_Gmail(gmail_ID, gmail_PW)
            if result_LogGM == True:
                break

        Regi_TID = pd.read_excel("특화서비스 등록 TID.xlsx")  # 입력 TID 리스트 파일
        # 메인 초대 진입점
        loof_Val = len(Regi_TID['TID'])
        for i in range(loof_Val):
            driver_Main.maximize_window()
            try_number = i + 1
            print("{}/{}번째 시작입니다.".format(i+1, loof_Val))

            # 데이터 추출 부분(From Exel)
            cur_TID = Regi_TID['TID'][i]
            cur_NAME = Regi_TID['NAME'][i]
            cur_PW = Regi_TID['PASSWORD'][i]
            cur_EMAIL = Regi_TID['EMAIL'][i]
            cur_PHONE = Regi_TID['PHONE'][i]

            # Gmail Inbox 비우기
            repeat_DelCount = 0
            while 1:
                if repeat_DelCount >= 10:
                    while 1:
                        imap_Gmail, result_LogGM = login_Gmail(
                            gmail_ID, gmail_PW)
                        if result_LogGM == True:
                            break
                imap_Gmail, result_Empty = empty_Gmail(imap_Gmail)
                if result_Empty == True:
                    repeat_DelCount = 0
                    break

            # NUGU DEV 대상자 중복초대 방지 삭제
            while 1:
                driver_Main, result_Delete = deleteUser_Nugu(
                    driver_Main, cur_TID)
                if result_Delete == True:
                    break

            # NUGU DEV 대상자 초대요청 보내기
            while 1:
                driver_Main, result_Invite = inviteUser_Nugu(
                    driver_Main, cur_TID, cur_NAME, cur_EMAIL, cur_PHONE)
                if result_Invite == True:
                    break
            driver_Main.minimize_window()

            # Gmail 초대 Link 찾기
            invite_URL = ''
            while 1:
                imap_Gmail, invite_URL = inviteLink_Gmail(imap_Gmail)
                if invite_URL != '':
                    break
                else:
                    count_Nomail += 1
                    if count_Nomail == 10:
                        count_Nomail = 0
                        result_NoMail = True
                        break

            # 초대 수락하기
            while 1:
                if result_NoMail == True:
                    print(cur_TID+"초대메일 찾을수 없음")
                    result_NoMail = False
                    break
                result_Invite = AcceptInvite_Nugu(invite_URL, cur_TID, cur_PW)
                if result_Invite == True:
                    break
                else:
                    print("초대 실패 재실행 시작")
    except:
        print("예외상황 프로그램 종료")
        try:
            driver_Main.close()
            imap_Gmail.logout()
        except:
            pass
    finally:
        print("프로그램 종료 최종 분기점")
        try:
            driver_Main.close()
            imap_Gmail.logout()
        except:
            pass

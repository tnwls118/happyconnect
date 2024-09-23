import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
import os

# SMTP 서버 설정
smtp_server = 'smtp.gmail.com'
smtp_port = 587
email_address = 'tnrruddl1220@gmail.com'
email_password = 'uapm mhhf eiqe aejb'

# 엑셀 파일 열기
wb = openpyxl.load_workbook(r"C:\Users\82109\Desktop\emails.xlsx")
sheet = wb.active

# SMTP 서버 연결 및 로그인
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email_address, email_password)

# 엑셀 파일에서 메일 정보 읽기
for row in sheet.iter_rows(min_row=2, values_only=True):
    recipient_email = row[0]   # 수신자 이메일
    subject = row[1]           # 메일 제목
    body = row[2]              # 메일 본문
    attachment_path = row[3]   # 첨부 파일 경로 (없을 수도 있음)

    # 이메일 메시지 설정
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # 메일 본문 추가
    msg.attach(MIMEText(body, 'plain'))

    # 첨부 파일이 있을 경우 파일 첨부
    if attachment_path:
        try:
            # 파일이 실제로 존재하는지 확인
            if os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    filename = os.path.basename(attachment_path)
                    part.add_header(
                        'Content-Disposition', f"attachment; filename={filename}")
                    msg.attach(part)
            else:
                print(f"파일이 존재하지 않음: {attachment_path}")
        except Exception as e:
            print(f"Error attaching file {attachment_path}: {e}")

    # 메일 전송
    server.sendmail(email_address, recipient_email, msg.as_string())
    print(f"메일 발송 완료: {recipient_email}")

# SMTP 서버 종료
server.quit()
print("모든 메일이 성공적으로 발송되었습니다.")

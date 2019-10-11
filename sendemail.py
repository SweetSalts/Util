import smtplib
from email.header import Header
from email.mime.text import MIMEText


def send_mail():
    smtpserver = 'mail.tcl.com'
    sender = 'siwei.yan@tcl.com'
    receivers = ['ysiwei97@163.com', 'zheng4.he@tcl.com']
    username = sender
    password = 'wuzi192514#'

    message = MIMEText('邮件发送测试......', 'plain', 'utf-8')
    message['From'] = Header(sender, 'utf-8')
    message['To'] = Header(';'.join(receivers), 'utf-8')
    subject = '邮件测试\n'
    message['Subject'] = Header(subject, 'utf-8')

    print(subject.strip('\n'))
    print('666')
    # try:
    #     smtp = smtplib.SMTP()
    #     smtp.connect(smtpserver, 25)
    #     smtp.login(username, password)
    #     smtp.sendmail(sender, receivers, message.as_string())
    #     smtp.quit()
    #     print('邮件发送成功' + receivers)
    # except smtplib.SMTPException as e:
    #     print('邮件发送失败, 失败原因:' + str(e))


def main():
    send_mail()


if __name__ == '__main__':
    main()

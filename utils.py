import random
import pandas as pd
from smtplib import SMTP
from datetime import date
from email.mime.text import MIMEText
from pyad import pyad_setdefaults, pyad
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def reset_user_password(user):
    try:
        pyad.set_defaults(ldap_server="DOMAIN.LOCAL", username="DOMAIN\ACCOUNTNAME",
                          password="SECURELY_STORED_PASSWORD")

        password = password_generator()

        user = pyad.aduser.ADUser.from_cn(user)
        user.set_password(f"{password}")

        print(f"RESET {user}")
        return password

    except Exception as e:
        print("Failed to reset password")
        print(e)
        exit()


def password_generator():
    df = pd.read_csv('Dictionary.csv')
    dictionary = df['Words']
    complies = False
    word1 = random.choice(dictionary)
    word2 = random.choice(dictionary)
    word3 = random.choice(dictionary)

    password = f'{word1} {word2} {word3}'

    while not complies:
        if len(password) > 45 or len(password) <= 35:
            password = f'{random.choice(dictionary)} {random.choice(dictionary)} {random.choice(dictionary)}'
            print(password)
            print(len(password))
        else:
            complies = True
            print(f"{password} COMPLIES WITH PASSWORD POLICY")
            print(len(password))

    return password


def send_log(users):
    msg = MIMEMultipart('related')
    msg['From'] = "SERVICE EMAIL ADDRESS GOES HERE"
    msg['Subject'] = f"New Hire Letter Generation Log for {date.today()}"

    if len(users) != 0:
        response = "Letters have been generated for the following users:<br>"
        breaks = "<br><br>"
    else:
        response = ""
        breaks = ""

    new_line = '<br>'
    msg_body = MIMEText("<p>Hello there,</p>"
                        "<p>Attached is the log for the programs last run.</p>"
                        f"<p>{response}{(new_line.join(users))}{breaks}"
                        "<br>Please report any strange behavior to <strong>S3cBar0n</strong>.</p>"
                        "<p>Thanks,<br>DocxFiller</p>", "html")
    msg.attach(msg_body)

    path = "LOGGING LOCATION PATH"
    log = f"{date.today()}.log"
    with open(path + log, "rb") as file:
        part = MIMEApplication(
            file.read(),
            Name=log
        )
    part['Content-Disposition'] = 'attachment; filename="%s"' % log
    msg.attach(part)

    smtp_server = SMTP("SMTP_IP:PORT")

    try:
        receiver = "EMAIL ADDRESS TO EMAIL THE LOG"
        msg['To'] = receiver
        smtp_server.sendmail(msg["From"], receiver, msg.as_string())
    except Exception as e:
        print(f"Failed to send post run email: {e}")

    smtp_server.quit()

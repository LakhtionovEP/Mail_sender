import smtplib
import email
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import os
import openpyxl
import random
import time
import sys
import datetime
import imaplib

#загрузка адресатов

wb = openpyxl.load_workbook('email_input.xlsx')
sheet = wb.active
recipients, names = [], []
i = 2
start_point = None
while sheet['A' + str(i)].value is not None:
    recipients.append(sheet['A' + str(i)].value)
    names.append(sheet['B' + str(i)].value)
    if sheet['C' + str(i)].value is None and start_point is None:
        start_point = i
    i += 1
if start_point is None:
    start_point = 2
count_work = start_point - 2
print(f'Обнаружено {len(recipients)} записей, из них {len(recipients) - (len(recipients))-count_work} необработанных, '
      f'нажмите 1 если хотите продолжить?')
while True:
    user_c = input()
    if user_c == '1':
        break
    else:
        sys.exit()
while True:
    print('Сколько адресатов нужно обработать?')
    user_q = int(input())
    if user_q > (len(recipients) - (len(recipients)-count_work)):
        print('Введеное значение больше количества необработанных адресатов')
        continue
    else:
        break

#Добавление вложения

filepath = r"Presentation.pdf"
basename = os.path.basename(filepath)

part = MIMEBase('application', "octet-stream")
part.set_payload(open(filepath, "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="%s"' % basename)

#Подключение к почтовому серверу

smtpObj = smtplib.SMTP('smtp.yandex.ru', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login('bms@uphill-consult.ru','***')

mail = imaplib.IMAP4_SSL('imap.yandex.ru')
mail.login('bms@uphill-consult.ru','***')

mail.list()
mail.select("inbox")

result, data = mail.search(None, "ALL")

ids = data[0]
id_list = ids.split()
latest_email_id = id_list[-1]

data = mail.fetch(latest_email_id, "(RFC822)")[1]
raw_email = data[0][1]
raw_email_string = raw_email.decode('utf-8')


email_message = email.message_from_string(raw_email_string)
sender = email_message['From']

t = time.time()
c, s, u = 0, 0, 0

#Набор из двух текстовок, которые рандомизированным образом подставляются в основное тело письма

text1 = """
Прошу Вас рассмотреть возможность сотрудничества в рамках проведения оценки различного вида имущества для нужд процессов по банкротству.

Надеюсь моё письмо Вас заинтересует, так как мы опытные специалисты и будем рады быть полезными в вашей работе. Ниже я напишу только основные моменты, которые нас выделяют среди наших коллег:
- Мы умеем слышать и всегда гибко и индивидуально подходим к решению каждой задачи.
- Нам более 10 лет, а в штате опытнейшие оценщики со всеми тремя квалификационными аттестатами, а ряд из них кандидаты экономических наук.
- Нам доверяют банки, государственные структуры, а также наши оценщики на постоянной основе привлекаются судьями (Московского городского суда, Арбитражного суда г. Москвы) в качестве экспертов для проведения финансово-экономической и оценочной экспертизы.
- При этом стоимость наших услуг и сроки исполнения конкурентны.

Во вложении я прикладываю презентацию о нашей экспертной организации.

Будем рады быть полезными.
"""
text2 = """
«Апхилл» - эксперты в области оценки. Просим Вас рассмотреть возможность сотрудничества в рамках проведения оценки различного вида имущества для нужд процессов, связанных с банкротством.


Надеюсь наше письмо Вас заинтересует, и мы будем полезны в вашей работе. 
Основные наши преимущества:
- Мы умеем слышать и всегда гибко и индивидуально подходим к решению каждой задачи.
- Нам более 10 лет, а в штате опытнейшие оценщики со всеми тремя квалификационными аттестатами, а ряд из них кандидаты экономических наук.
- Нам доверяют банки, государственные структуры, а также наши оценщики на постоянной основе привлекаются судьями (Московского городского суда, Арбитражного суда г. Москвы) в качестве экспертов для проведения финансово-экономической и оценочной экспертизы.
- При этом стоимость наших услуг и сроки исполнения конкурентны.

Во вложении прикладываем презентацию о нашей экспертной организации.

Надеемся на ваше положительное решение.
"""

#Компоновка итогового письма

for x in range(start_point-2, start_point - 2 + user_q):
    recipient = recipients[x]
    name = names[x]
    rest = random.randint(8, 15) + random.random()
    text_body = random.choice((text1, text2))
    msg = MIMEMultipart()
    text = MIMEText(f"{name}, Добрый день!\n {text_body}", 'plain', 'utf-8')
    sign = MIMEText(
        """
С уважением,
 
Старший консультант

 Консалтинговая группа «Апхилл»
107031, г. Москва, ул. Кузнецкий Мост, д. 19, стр. 1
www.uphill.ru"""
        , 'plain', 'utf-8')

    msg['Subject'] = Header("Предложение о сотрудничестве от независимой оценочной компании «Апхилл»", 'utf-8')
    msg['From'] = "bms@uphill-consult.ru"
    msg['To'] = f"{recipient}"
    msg.attach(text)
    msg.attach(part)
    msg.attach(sign)

#Статистика по проведенной работе

    try:
        smtpObj.sendmail('bms@uphill-consult.ru', f'{recipient}', msg.as_string())
        sheet['C' + str(x + 2)] = 'Success'
        s += 1
    except:
        sheet['C' + str(x + 2)] = 'Failed'
        u += 1
    finally:
        sheet['D' + str(x + 2)] = datetime.date.today().strftime('%d.%m.%Y')
        sys.stdout.write('\r%s' % c + ' из ' + str(user_q) + ' адресов уже обработано ')
        sys.stdout.write(str(int((time.time() - t) // 60)) + ' минут ' +
                         str(float('{:.1f}'.format((time.time() - t) % 60))) + ' секунд затрачено')
        sys.stdout.flush()
        c += 1
        time.sleep(5)
        if recipient in str(email_message) and sender == 'mailer-daemon@yandex.ru':
            sheet['C' + str(x + 2)] = 'Error'
            s -= 1
            u += 1
        time.sleep(rest)
       # print(rest)
print('\nОтправка завершена!')
print('Обработано адресатов: ' + str(user_q))
print('Успешно отправлено: ' + str(s))
print('Не отправлено по ошибке: ' + str(u))
print('Общее время: ' + (str(int((time.time() - t) // 60)) + ' минут ' + str(int(time.time() - t) % 60)) + ' секунд')
wb.save('email_input.xlsx')
smtpObj.quit()
os.system('pause')
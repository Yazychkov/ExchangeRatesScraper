import time
import settings
from selenium import webdriver
import xlsxwriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import os


class Parser:
    def __init__(self):
        self.url = "https://yandex.ru/"
        self.driver = webdriver.Edge("msedgedriver.exe")
        self.wb = xlsxwriter.Workbook('test.xlsx', {'strings_to_numbers': True})
        self.ws = self.wb.add_worksheet()
        self.res_usd = list()
        self.res_euro = list()
        self.usd_list_for_excel = list()
        self.euro_list_for_excel = list()
        self.coef = list()
        self.len_excel = int()
        self.number_str = int()

        self.parsing_data()
        self.prepare_data()
        self.load_data_to_excel()

    def parsing_data(self):
        try:
            ### Открытие страницы Яндекса, добавляем в список items элементы классов (USD, EURO, нефть)
            self.driver.get(self.url)
            time.sleep(5)
            items = self.driver.find_elements_by_xpath('//a[@class="home-link home-link_black_yes inline-stocks__link"]')
            items[0].click()
            time.sleep(5)

            ### Открытие первой вкладки с USD, добавляем в список res_usd элементы таблицы курса
            self.driver.switch_to.window(self.driver.window_handles[1])
            time.sleep(5)
            usd_course = self.driver.find_elements_by_class_name('news-stock-table__cell')
            for elem in usd_course:
                self.res_usd.append(elem.text)
            time.sleep(5)

            ### Закрываем вкладку, возвращаемся на 0 вкладку Яндекса и переходим на вкладку с euro
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(5)
            items[1].click()

            ### Парсим данные с euro
            self.driver.switch_to.window(self.driver.window_handles[1])
            time.sleep(5)
            euro_course = self.driver.find_elements_by_class_name('news-stock-table__cell')
            for elem in euro_course:
                self.res_euro.append(elem.text)
            time.sleep(5)
        except Exception:
            print('Возможно Вы слишком часто запрашивали данные, повторите попытку позже')
        finally:
            self.driver.close()
            self.driver.quit()

    def prepare_data(self):
        usd_course_list = list()
        usd_coef_list = list()

        for usd_course in self.res_usd[1::3]:
            if usd_course == self.res_usd[1]:
                usd_course_list.append(usd_course)
                continue
            usd_course_list.append(float(usd_course.replace(',', '.')))

        for usd_coef in self.res_usd[2::3]:
            if usd_coef == self.res_usd[2]:
                usd_coef_list.append(usd_coef)
                continue
            usd_coef_list.append(float(usd_coef.replace(',', '.')))
        self.usd_list_for_excel.append(self.res_usd[::3])
        self.usd_list_for_excel.append(usd_course_list)
        self.usd_list_for_excel.append(usd_coef_list)

        euro_course_list = list()
        euro_coef_list = list()

        for euro_course in self.res_euro[1::3]:
            if euro_course == self.res_euro[1]:
                euro_course_list.append(euro_course)
                continue
            euro_course_list.append(float(euro_course.replace(',', '.')))

        for euro_coef in self.res_euro[2::3]:
            if euro_coef == self.res_euro[2]:
                euro_coef_list.append(euro_coef)
                continue
            euro_coef_list.append(float(euro_coef.replace(',', '.')))
        self.euro_list_for_excel.append(self.res_euro[::3])
        self.euro_list_for_excel.append(euro_course_list)
        self.euro_list_for_excel.append(euro_coef_list)

        ### Подготавливаем данные сотношение евро к доллару
        euro_list_for_coef = []
        usd_list_for_coef = []
        for euro in self.res_euro[4::3]:
            euro_list_for_coef.append(float(euro.replace(',', '.')))
        for usd in self.res_usd[4::3]:
            usd_list_for_coef.append(float(usd.replace(',', '.')))
        self.coef = [round(e / d, 4) for e, d in zip(euro_list_for_coef, usd_list_for_coef)]

        ### Подготавливаем данные для письма
        len_excel = len(self.euro_list_for_excel[1])
        self.len_excel = int(len_excel)
        if int(str(len_excel)[-1]) == 1 and len_excel < 10  or len_excel > 19:
            self.number_str = 'строка'
        elif 4 >= int(str(self.len_excel)[-1]) >= 2 and len_excel < 10  or len_excel > 19:
            self.number_str = 'строки'
        else:
            self.number_str = 'строк'

    def load_data_to_excel(self):
        format0 = self.wb.add_format({'align': 'center', 'num_format': 'dd/mm/yy'})
        format1 = self.wb.add_format({'align': 'center', 'num_format': '$#,##0.00'})
        format1_2 = self.wb.add_format({'align': 'center', 'num_format': '€#,##0.00'})
        format2 = self.wb.add_format({'align': 'center'})

        self.ws.write_column('A1', self.usd_list_for_excel[0], format0)
        self.ws.write_column('B1', self.usd_list_for_excel[1], format1)
        self.ws.write_column('C1', self.usd_list_for_excel[2], format2)

        self.ws.write_column('D1', self.euro_list_for_excel[0], format0)
        self.ws.write_column('E1', self.euro_list_for_excel[1], format1_2)
        self.ws.write_column('F1', self.euro_list_for_excel[2], format2)

        self.ws.write_column('G2', self.coef, format2)

        self.wb.close()

    def send_message(self):
        server = settings.server
        user = settings.user_email
        password = settings.user_password
        recipient = settings.recipient
        subject = 'Курс валют'
        text = 'Курс валют за последние 10 дней, в экселе {} {}'.format(self.len_excel, self.number_str)

        filepath = "test.xlsx"
        basename = os.path.basename(filepath)
        filesize = os.path.getsize(filepath)

        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = user
        msg['To'] = recipient

        text = MIMEText(text, 'plain')
        attachment = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
        attachment.set_payload(open(filepath, "rb").read())
        attachment.add_header('Content-Description', basename)
        attachment.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'.format(basename, filesize))
        encoders.encode_base64(attachment)

        msg.attach(text)
        msg.attach(attachment)

        mail = smtplib.SMTP_SSL(server)
        mail.login(user, password)
        mail.sendmail(user, recipient, msg.as_string())
        mail.quit()


if __name__ == "__main__":
    exp = Parser()
    print(exp.send_message())

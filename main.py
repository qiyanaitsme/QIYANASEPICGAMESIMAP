import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, QSettings, Qt

import imaplib
import email
from bs4 import BeautifulSoup
import logging
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class EmailClient:
    def __init__(self, host, port, username=None, password=None):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.mail = None
        self.connected = False

    def connect(self):
        try:
            self.mail = imaplib.IMAP4_SSL(self.host, self.port)
            self.mail.login(self.username, self.password)
            if not self.connected:
                logging.info("Успешное подключение к почтовому серверу.")
                self.connected = True
        except imaplib.IMAP4.error as e:
            logging.error(f"Ошибка подключения: {e}")
            self.connected = False

    def disconnect(self):
        try:
            if self.mail:
                self.mail.logout()
                logging.info("Отключение от почтового сервера.")
        except Exception as e:
            logging.error(f"Ошибка при отключении от почтового сервера: {e}")

    def change_credentials(self, new_username, new_password):
        self.username = new_username
        self.password = new_password

    def find_email_by_subject(self, subject):
        try:
            self.mail.select('inbox')
            _, data = self.mail.search(None, f'(SUBJECT "{subject}")')
            return data[0].split()
        except Exception as e:
            logging.error(f"Ошибка при поиске почты: {e}")
            return []

    def fetch_email_content(self, email_id):
        try:
            _, msg_data = self.mail.fetch(email_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])
            return msg
        except Exception as e:
            logging.error(f"Ошибка при получении содержимого письма: {e}")
            return None

    def extract_verification_codes(self, msg):
        try:
            verification_codes = []

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    payload = part.get_payload(decode=True)
                    soup = BeautifulSoup(payload, 'html.parser')

                    td_elements = soup.find_all('td', {'style': 'background:#f1f1f1;margin-top:20px;font-family: arial,helvetica,sans-serif; mso-line-height-rule: exactly; font-size:30px; color:#202020; line-height:19px; line-height: 134%; letter-spacing: 10px;text-align: center;padding: 20px 0px !important;letter-spacing: 10px !important;border-radius: 4px;'})

                    verification_codes.extend([td.get_text(strip=True).strip() for td in td_elements])

            return verification_codes
        except Exception as e:
            logging.error(f"Ошибка при извлечении кодов верификации: {e}")
            return []

class EmailClientApp(QWidget):
    def __init__(self):
        super().__init__()

        self.settings = QSettings('QiyanaINC', 'EpicGamesIMAP')

        self.init_ui()

        self.email_client = None
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.on_timer_timeout)

        self.load_settings()

    def init_ui(self):
        layout = QVBoxLayout()

        self.email_password_input = QLineEdit(self)
        self.email_password_input.setPlaceholderText('Введите email:password от почты')

        self.imap_input = QLineEdit(self)
        self.imap_input.setPlaceholderText('Введите адрес IMAP сервера')

        self.port_input = QLineEdit(self)
        self.port_input.setPlaceholderText('Введите порт IMAP сервера')

        self.get_email_button = QPushButton('Получить письмо', self)
        self.get_email_button.clicked.connect(self.on_get_email_button_clicked)

        layout.addWidget(self.email_password_input)
        layout.addWidget(self.imap_input)
        layout.addWidget(self.port_input)
        layout.addWidget(self.get_email_button)

        self.setLayout(layout)

        self.setWindowTitle('QIYANAS EPICGAMES IMAP')
        self.setFixedSize(280, 130)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, False)
        self.setWindowIcon(QIcon(resource_path('res/icon.ico')))

    def on_get_email_button_clicked(self):
        email_password = self.email_password_input.text()
        imap_address = self.imap_input.text()
        imap_port = self.port_input.text()

        try:
            email, password = email_password.split(':')
        except ValueError:
            QMessageBox.warning(self, 'Предупреждение', 'Введите email и пароль в формате email:password')
            return

        if not email or not password or not imap_address or not imap_port:
            QMessageBox.warning(self, 'Предупреждение', 'Введите все необходимые данные.')
            return

        self.email_client = EmailClient(
            host=imap_address,
            port=int(imap_port),
            username=email,
            password=password
        )

        self.email_client.connect()

        if self.email_client.connected:
            email_subject = "Epic Games - Email Verification"
            email_ids = self.email_client.find_email_by_subject(email_subject)

            if email_ids:
                email_id = email_ids[-1]
                msg = self.email_client.fetch_email_content(email_id)

                if msg:
                    verification_codes = self.email_client.extract_verification_codes(msg)

                    if verification_codes:
                        QMessageBox.information(self, 'Код верификации', f'Ваш код верификации: {", ".join(verification_codes)}')
                    else:
                        QMessageBox.warning(self, 'Предупреждение', 'Коды верификации не найдены в письме.')
                else:
                    QMessageBox.warning(self, 'Предупреждение', 'Ошибка при получении содержимого письма.')
            else:
                QMessageBox.warning(self, 'Предупреждение', f'Письмо с темой {email_subject} не найдено.')

            self.timer.start(30000)
        else:
            QMessageBox.critical(self, 'Ошибка', 'Ошибка подключения к почтовому серверу.')

    def on_timer_timeout(self):
        if self.email_client:
            self.email_client.disconnect()
            QMessageBox.warning(self, 'Таймаут', 'Соединение разорвано. Введите новые данные.')
            self.timer.stop()

    def load_settings(self):
        self.email_password_input.setText(self.settings.value('EmailPassword', ''))
        self.imap_input.setText(self.settings.value('ImapAddress', ''))
        self.port_input.setText(self.settings.value('ImapPort', ''))

    def save_settings(self):
        self.settings.setValue('EmailPassword', self.email_password_input.text())
        self.settings.setValue('ImapAddress', self.imap_input.text())
        self.settings.setValue('ImapPort', self.port_input.text())

    def closeEvent(self, event):
        self.save_settings()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailClientApp()
    window.show()
    sys.exit(app.exec_())
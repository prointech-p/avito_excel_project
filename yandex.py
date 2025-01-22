import imaplib
import email
import re
from email.header import decode_header
import json

settings_path = 'settings.json'


class EmailDownloader:
    def __init__(self, settings_file):
        self.settings = self.load_settings(settings_file)
        self.email_address = self.settings["Настройки"]["Почта_с_остатками"]
        self.password = self.settings["Настройки"]["Программный_пароль_к_почте_с_остатками"]
        self.imap_server = self.settings["Настройки"]["IMAP_SERVER"]
        self.new_file_name = self.settings["Настройки"]["Остатки"]
        # self.download_folder = download_folder
        self.mail = None

        # Создание папки для вложений, если не существует
        # if not os.path.exists(self.download_folder):
        #     os.makedirs(self.download_folder)

    def load_settings(self, settings_file):
        with open(settings_file, 'r', encoding='utf-8') as file:
            return json.load(file)  # Возвращаем весь содержимое файла

    def connect(self):
        """Подключение к IMAP серверу."""
        self.mail = imaplib.IMAP4_SSL(self.imap_server)
        self.mail.login(self.email_address, self.password)

    def fetch_emails(self):
        """Получение и обработка писем."""
        self.mail.select('inbox')
        status, messages = self.mail.search(None, 'ALL')
        mail_ids = messages[0].split()

        latest_attachment = None

        for mail_id in mail_ids:
            status, msg_data = self.mail.fetch(mail_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            subject = msg['subject']
            print(f'Subject: {subject}')

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        filename = part.get_filename()
                        if filename:
                            decoded_filename = self.decode_mime_words(filename)
                            if self.is_matching_filename(decoded_filename):
                                latest_attachment = part

        if latest_attachment:
            self.download_attachment(latest_attachment)

    def is_matching_filename(self, filename):
        """Проверка, соответствует ли имя файла требуемому шаблону."""
        pattern = r'^Остатки на \d{2}-\d{2}-\d{4} \d{2}-\d{2}-\d{2}\.xlsx$'
        return re.match(pattern, filename) is not None

    def download_attachment(self, part):
        """Скачивание вложения."""
        filename = part.get_filename()
        if filename:
            decoded_filename = self.decode_mime_words(filename)
            # filepath = os.path.join(self.download_folder, decoded_filename)
            filepath = self.new_file_name
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f'Скачано: {decoded_filename}')

    def decode_mime_words(self, raw_string):
        """Декодирование закодированной строки в формате MIME."""
        decoded_fragments = decode_header(raw_string)
        decoded_string = ''
        for fragment, encoding in decoded_fragments:
            if isinstance(fragment, bytes):
                fragment = fragment.decode(encoding or 'utf-8')
            decoded_string += fragment
        return decoded_string

    def logout(self):
        """Закрытие соединения с почтовым сервером."""
        self.mail.logout()


if __name__ == '__main__':
    email_downloader = EmailDownloader(settings_path)
    email_downloader.connect()
    email_downloader.fetch_emails()
    email_downloader.logout()

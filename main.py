import argparse
import tkinter as tk
from price import AvitoFilesHandler
from avito_app import App
from yandex import EmailDownloader


settings_path = 'settings.json'  # Укажите путь к вашему файлу настроек


def main():
    parser = argparse.ArgumentParser(description="Запуск различных модулей.")
    parser.add_argument("command", choices=["automation", "app"], help="команда для выполнения")

    args = parser.parse_args()

    if args.command == "automation":
        price_handler = AvitoFilesHandler(settings_path)
        price_handler.create_price_file()
        price_handler.create_all_avito_files(only_active=True)
        price_handler.create_all_avito_files(only_active=False)
    elif args.command == "app":
        root = tk.Tk()
        app = App(root)
        root.mainloop()


# Пример использования
if __name__ == "__main__":
    main()


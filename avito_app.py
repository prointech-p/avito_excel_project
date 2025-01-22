import tkinter as tk
from tkinter import messagebox
from price import AvitoFilesHandler

SETTINGS_FILE = 'settings.json'
YELLOW = "#f7f5dd"
GREEN = "#9bdeac"


class App:
    def __init__(self, master):
        self.master = master
        master.title("Авито")
        master.config(padx=50, pady=25, bg=YELLOW)

        # Установка размеров окна (ширина x высота)
        window_width = 500
        window_height = 400

        # Получаем размеры экрана
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()

        # Вычисляем координаты для центрирования окна
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        # Устанавливаем размеры и положение окна
        master.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Заголовок
        self.label = tk.Label(master, text="Обновление", font=("Arial", 16), bg=YELLOW)
        self.label.grid(row=0, column=0, columnspan=2, pady=10)

        # Кнопка "Прайс.xlsx"
        self.price_button = tk.Button(master, text="Прайс.xlsx", command=self.create_price_file, width=20,
                                      font=("Arial", 14), bg=GREEN)
        self.price_button.grid(row=1, column=0, pady=5, sticky="w")

        # Чекбокс "Проверять почту"
        self.check_email_var = tk.BooleanVar(value=False)
        self.check_email_checkbox = tk.Checkbutton(master, text="Проверять почту", variable=self.check_email_var,
                                                   font=("Arial", 14), bg=YELLOW)
        self.check_email_checkbox.grid(row=1, column=1, pady=5, sticky="w")

        # Кнопка "Заказы", объединяющая 2 строки
        self.orders_button = tk.Button(master, text="Заказы", command=self.create_orders_files, width=20, height=2,
                                       font=("Arial", 14), bg=GREEN)
        self.orders_button.grid(row=2, column=0, rowspan=2, pady=5, sticky="w")

        # Чекбоксы справа от "Заказы", расположенные вертикально
        self.orders02_checkbox_var = tk.BooleanVar(value=True)
        self.orders02_checkbox = tk.Checkbutton(master, text="02 Краснодар", variable=self.orders02_checkbox_var,
                                        font=("Arial", 13), bg=YELLOW)
        self.orders02_checkbox.grid(row=2, column=1, sticky="w", pady=(5, 0))

        self.orders03_checkbox_var = tk.BooleanVar(value=True)
        self.orders03_checkbox = tk.Checkbutton(master, text="03 Ростов", variable=self.orders03_checkbox_var,
                                        font=("Arial", 13), bg=YELLOW)
        self.orders03_checkbox.grid(row=3, column=1, sticky="w", pady=(0, 5))

        # Кнопка "Авито.xlsx"
        self.avito_button = tk.Button(master, text="Авито.xlsx", command=self.create_avito_files, width=20,
                                      font=("Arial", 14), bg=GREEN)
        self.avito_button.grid(row=4, column=0, pady=5, sticky="w")

        # Чекбокс "Только активные"
        self.only_active_var = tk.BooleanVar(value=True)
        self.active_checkbox = tk.Checkbutton(master, text="Только активные", variable=self.only_active_var,
                                              font=("Arial", 14), bg=YELLOW)
        self.active_checkbox.grid(row=4, column=1, pady=5, sticky="w")

        # Кнопка "Цены и вес"
        self.prices_weights_button = tk.Button(master, text="Цены и вес", command=self.show_prices_weights, width=20,
                                               font=("Arial", 14), bg=GREEN)
        self.prices_weights_button.grid(row=5, column=0, pady=5, sticky="w")


    def create_price_file(self):
        check_email = self.check_email_var.get()
        price_handler = AvitoFilesHandler(SETTINGS_FILE)
        price_handler.create_price_file(check_email)
        messagebox.showinfo("Уведомление", "Файл 'Прайс.xlsx' создан.")

    def create_orders_files(self):
        check_email = self.check_email_var.get()
        check_02 = self.orders02_checkbox_var.get()
        check_03 = self.orders03_checkbox_var.get()
        stores = []
        if check_02:
            stores.append("02")
        if check_03:
            stores.append("03")
        if stores:
            price_handler = AvitoFilesHandler(SETTINGS_FILE)
            price_handler.create_orders_files(stores=stores, check_email=check_email)
            messagebox.showinfo("Уведомление", "Файлы заказов созданы.")
        else:
            messagebox.showinfo("Уведомление", "Не выбран ни один из магазинов")

    def create_avito_files(self):
        only_active = self.only_active_var.get()
        price_handler = AvitoFilesHandler(SETTINGS_FILE)
        price_handler.create_all_avito_files(only_active)
        messagebox.showinfo("Уведомление", "Файлы 'Авито.xlsx' созданы.")

    def show_prices_weights(self):
        # "Цены и вес"
        price_handler = AvitoFilesHandler(SETTINGS_FILE)
        price_handler.create_price_weigh_file()
        messagebox.showinfo("Уведомление", "Обработка 'Цены и вес' выполнена.")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

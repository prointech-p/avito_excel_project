import requests
from pprint import pprint

def get_warehouses_data(article, api_key):
    """
    Делает запрос к API, извлекает и возвращает данные из ключа 'warehouses'.

    :param article: Артикул продукта для запроса.
    :param api_key: API-ключ пользователя.
    :return: Словарь с данными 'warehouses' или пустой словарь, если ключ отсутствует.
    """
    url = f"https://b2b.autofamily.ru/api/getProduct/?article={article}&key={api_key}&v=1.1"
    try:
        # Отправка GET-запроса
        response = requests.get(url)
        response.raise_for_status()  # Проверка на успешность ответа
        print(response)

        # Разбор JSON-ответа
        data = response.json()

         # Извлечение ключа 'response'
        response_data = data.get("response", {})
        pprint(response_data)

        name = response_data.get("name", "")

        # Проверка наличия ключа 'warehouses' в 'response'
        warehouses = response_data.get("warehouses", {})
        if not isinstance(warehouses, dict):
            raise ValueError("'warehouses' должен быть словарём")
        
        result = {"warehouses": warehouses, "name": name}
        
        return result

    except requests.exceptions.RequestException as e:
        print(f"Ошибка соединения с API: {e}")
    except ValueError as e:
        print(f"Ошибка обработки данных: {e}")
    except Exception as e:
        print(f"Непредвиденная ошибка: {e}")
    
    # Возвращаем пустой словарь, если произошла ошибка
    return {}


def test_price_format(self):
        price_path = self.settings["Настройки"]["Прайс_для_работы"]

        # Создаем новую книгу, если прайс еще не существует
        try:
            workbook = load_workbook(
                filename=price_path, read_only=False, keep_vba=True)
        except FileNotFoundError:
            workbook = Workbook()

        sheet = workbook.active
        sheet.title = "Прайс"

        # Переименовываем столбцы 7 и 8
        sheet.cell(row=1, column=7).value = "Цена"
        sheet.cell(row=1, column=8).value = "Расп"

        # Получаем настройки аббревиатур городов
        abbreviations = self.settings["Аббревиатуры городов"]

        # Проходим по всем ячейкам в первой строке
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            # Проверяем, есть ли значение в настройках
            if cell_value in abbreviations:
                # Заменяем значение на аббревиатуру
                sheet.cell(
                    row=1, column=col).value = abbreviations[cell_value]["Аббревиатура"]
                sheet.column_dimensions[sheet.cell(
                    row=1, column=col).column_letter].width = 7  # Уменьшаем размер
                sheet.cell(row=row, column=col).alignment = Alignment(horizontal='center')

        # Сортируем города по их "Коду"
        sorted_cities = sorted(abbreviations.items(),
                               key=lambda x: int(x[1]["Сортировка"]))

        # Определяем, на каком столбце будем добавлять новые
        start_col = sheet.max_column + 1
        cols_to_delete = []
        # Добавляем столбцы для каждого города
        for index, (city, data) in enumerate(sorted_cities):
            column_letter = start_col + index
            cell = sheet.cell(row=1, column=column_letter,
                              value="Аббревиатура")
            cell.value = data["Аббревиатура"]
            new_color = data["Цвет"]

            # Проходим по первой строке до start_col и ищем совпадения
            for col in range(1, start_col):
                if sheet.cell(row=1, column=col).value == data["Аббревиатура"]:
                    cols_to_delete.append(col)
                    # Копируем значения из найденного столбца
                    for row in range(2, sheet.max_row + 1):
                        sheet.cell(row=row, column=column_letter).value = sheet.cell(
                            row=row, column=col).value
                        sheet.cell(row=row, column=column_letter).fill = PatternFill(start_color=new_color,
                                                                                     end_color=new_color,
                                                                                     fill_type='solid')
                        sheet.cell(row=row, column=column_letter).alignment = Alignment(horizontal='center')
        # Опять проходим по первой строке до start_col и удаляем старые столбцы городов
        cols_to_delete.sort(reverse=True)
        # print(cols_to_delete)
        for col in cols_to_delete:
            sheet.delete_cols(col, 1)  # Удаляем столбец

        # Получаем номер последнего столбца
        last_col = sheet.max_column

        # Копируем столбцы 3-6 в конец таблицы
        for col in range(3, 7):  # Столбцы 3, 4, 5, 6
            for row in range(1, sheet.max_row + 1):
                sheet.cell(row=row, column=last_col + (col - 3) +
                           1).value = sheet.cell(row=row, column=col).value

        # Удаляем исходные столбцы 3-6
        sheet.delete_cols(3, 4)  # Удаляем 4 столбца, начиная с 3-го

        # Добавляем столбцы: 1 - "Вес", 2 и 3 - пустые, 4 - "ID_01"
        sheet.insert_cols(1, 5)  # Вставляем 4 столбца (1, 2, 3, 4)
        sheet.cell(row=1, column=1, value="Вес")
        sheet.cell(row=1, column=2, value="")
        sheet.cell(row=1, column=3, value="")
        sheet.cell(row=1, column=4, value="")
        sheet.cell(row=1, column=5, value="ID_01")

        # Устанавливаем ширину столбцов
        sheet.column_dimensions[sheet.cell(
            row=1, column=1).column_letter].width = 4.17  # 1-й столбец (Вес)
        sheet.column_dimensions[sheet.cell(
            row=1, column=2).column_letter].width = 0.82  # 2-й столбец
        sheet.column_dimensions[sheet.cell(
            row=1, column=3).column_letter].width = 0.82  # 3-й столбец
        sheet.column_dimensions[sheet.cell(
            row=1, column=4).column_letter].width = pixels_to_width(9)  # 4-й столбец
        sheet.column_dimensions[sheet.cell(
            row=1, column=5).column_letter].hidden = True  # 5-й столбец (ID_01)

        # Добавляем столбцы с параметрами из раздела "Авито"
        avito_settings = self.settings["Авито"]  # Обращаемся к разделу "Авито"
        # Начинаем с 6-го столбца (5 - потому что будем сразу увеличивать на 1)
        col_n = 5

        for key, value in avito_settings.items():
            col_n += 1  # Увеличиваем смещение для следующего столбца
            column_name = f"Avitoid_{value['Код']}"
            # Вставляем новый столбец перед добавлением заголовка
            sheet.insert_cols(col_n)
            # Заполняем заголовок
            sheet.cell(row=1, column=col_n, value=column_name)

        # Устанавливаем ширину для дополнительных столбцов
        for i in range(5, col_n + 1):  # 5-й столбец и дальше
            # Ширина по вашему выбору
            # sheet.column_dimensions[sheet.cell(
            #     row=1, column=i).column_letter].width = 11
            sheet.column_dimensions[sheet.cell(
                row=1, column=i).column_letter].hidden = True

        col_n += 1
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = pixels_to_width(145)  # Размер для Артикул
        col_n += 1
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = pixels_to_width(790)  # Размер для Товар
        col_n += 1

        # Создаем словарь кодов городов, где ключ — Код, значение — Аббревиатура
        city_params_dict = {}
        city_code_dict = {}
        city_api_dict = {}
        for city, details in abbreviations.items():
            code = details["Код"]
            abbreviation = details["Аббревиатура"]
            api_key = details["Обновление по API"]
            city_code_dict[code] = abbreviation
            city_api_dict[abbreviation] = api_key
        city_params_dict['codes'] = city_code_dict
        city_params_dict['api_keys'] = city_api_dict

        # Преобразование словаря в строку JSON
        json_str = json.dumps(city_params_dict)

        sheet.insert_cols(col_n)  # Добавляем пустой столбец
        # sheet.cell(row=1, column=col_n, value="")
        # В заголовок записываем справочник кодов городов
        sheet.cell(row=1, column=col_n).value = json_str
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = 2

        col_n += 1
        # Вставляем новый столбец перед добавлением заголовка
        sheet.insert_cols(col_n)
        sheet.cell(row=1, column=col_n, value="Прайс")  # Заполняем заголовок
        # Заливаем столбец Прайс цветом (242,242,242)
        for row in range(1, sheet.max_row + 1):
            sheet.cell(row=row, column=col_n).fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2',
                                                                 fill_type='solid')
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = 10  # Размер для Прайс
        col_n += 1
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = 10  # Размер для Цена
        
        col_n += 1
        # Вставляем новый столбец -10% перед добавлением заголовка
        sheet.insert_cols(col_n)
        sheet.cell(row=1, column=col_n, value="-10%")  # Заполняем заголовок
        for row in range(2, sheet.max_row + 1):
            price_str = sheet.cell(row=row, column=col_n-1).value
            price_val = 0
            try:
                if price_str is not None:  # Проверяем, что ячейка не пуста
                    price_val = float(price_str)
            except ValueError:
                pass
            price_10_val = round(price_val * 0.9, 2)
            sheet.cell(row=row, column=col_n).value = price_10_val
        sheet.column_dimensions[get_column_letter(col_n)].hidden = True  # Скрываем столбец -10%

        col_n += 1
        sheet.column_dimensions[sheet.cell(
            row=1, column=col_n).column_letter].width = 10  # Размер для Расп
        # col_n += 1
        # sheet.insert_cols(col_n)  # Вставляем новый столбец перед добавлением заголовка

        # # Создаем словарь кодов городов, где ключ — Код, значение — Аббревиатура
        # city_params_dict = {}
        # city_code_dict = {}
        # city_api_dict = {}
        # for city, details in abbreviations.items():
        #     code = details["Код"]
        #     abbreviation = details["Аббревиатура"]
        #     api_key = details["Обновление по API"]
        #     city_code_dict[code] = abbreviation
        #     city_api_dict[abbreviation] = api_key
        # city_params_dict['codes'] = city_code_dict
        # city_params_dict['api_keys'] = city_api_dict

        # # Преобразование словаря в строку JSON
        # json_str = json.dumps(city_params_dict)

        # sheet.cell(row=1, column=col_n).value = json_str  # В заголовок записываем справочник кодов городов
        # sheet.column_dimensions[sheet.cell(row=1, column=col_n).column_letter].width = 2

        # Устанавливаем размер шрифта в первой строке
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=1, column=col).font = Font(size=10)

        # Заливаем первую строку цветом (79,129,189) и устанавливаем белый цвет шрифта
        header_fill = PatternFill(
            start_color='4F7DBE', end_color='4F7DBE', fill_type='solid')
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=1, column=col).fill = header_fill
            sheet.cell(row=1, column=col).font = Font(
                color='FFFFFF')  # Белый цвет

        # Устанавливаем сетку с цветом (79,129,189)
        border_color = Side(style='thin', color='4F7DBE')
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.border = Border(
                    left=border_color, right=border_color, top=border_color, bottom=border_color)

        # Замораживаем первую строку
        sheet.freeze_panes = sheet["A2"]

        # Сохраняем изменения
        workbook.save(price_path)
        print(f"Таблица прайса успешно сформирована и сохранена в '{
              price_path}'.")

# Пример использования
if __name__ == "__main__":
    article = "ELEMENT245807NOMK"  # Пример артикула
    api_key = "mMS6OL8RtP5Ca9iHCfvAIGUz5Qcd39fB"  # Ваш API-ключ
    warehouses_data = get_warehouses_data(article, api_key)

    if warehouses_data:
        print("Данные 'warehouses':")
        print(warehouses_data)
    else:
        print("Ключ 'warehouses' отсутствует или произошла ошибка.")




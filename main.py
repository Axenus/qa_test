from pywinauto import Application  # Импортируем класс Application из библиотеки pywinauto
import os  # Импортируем модуль os для работы с путями файловой системы
import pytest  # Импортируем pytest для запуска тестов

# Указываем базовый путь к каталогу с приложением
BASE_DIR = r'C:\\Users\\euene\\OneDrive\\Desktop\\headHunterTest-main'


def start_application():
    """
    Функция запускает приложение и возвращает объект приложения.
    """
    app = Application().start(os.path.join(BASE_DIR, 'main.exe'))  # Запускаем приложение и получаем объект приложения
    return app


def test_about_program():
    """
    Тест проверяет функциональность пункта меню «? -> О программе».
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения (по части заголовка окна)
    window.menu_select("?->О программе")  # Открываем пункт меню «? -> О программе»
    assert "Тестовое приложение для соискателей на должность QA." in window.static_texts()  # Проверяем наличие текста в окне
    app.kill()  # Закрываем приложение после выполнения теста


def test_activation_key():
    """
    Тест проверяет функциональность ввода и проверки ключа активации.
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения
    window.menu_select("?->О программе")  # Открываем пункт меню «? -> О программе»
    window.Edit.set_text("VALIDKEY")  # Вводим корректный ключ активации
    window.OK.click()  # Нажимаем кнопку "Ок"
    assert "Успешно!" in window.static_texts()  # Проверяем, что отображается сообщение об успешной активации
    window.Edit.set_text("INVALIDKEY")  # Вводим некорректный ключ активации
    window.OK.click()  # Нажимаем кнопку "Ок"
    assert "Ошибка" in window.static_texts()  # Проверяем, что отображается сообщение об ошибке
    app.kill()  # Закрываем приложение после выполнения теста


def test_convert_ed807_to_packetepd():
    """
    Тест проверяет функциональность конвертации файла ED807 в PacketEPD.
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения
    window.menu_select("Convert->ED807toPacketEPD")  # Открываем пункт меню «Convert -> ED807toPacketEPD»
    file_dialog = app.window(title_re=".*Open.*")  # Находим диалоговое окно "Открыть файл"
    file_dialog.Edit.set_text(os.path.join(BASE_DIR, 'test.ED807.xml'))  # Вводим путь к тестовому файлу ED807
    file_dialog.Open.click()  # Нажимаем кнопку "Открыть"
    assert "Конвертация проведена успешно" in window.static_texts()  # Проверяем, что отображается сообщение об успешной конвертации
    assert os.path.exists(os.path.join(BASE_DIR, 'PacketEPD.xml'))  # Проверяем, что файл PacketEPD.xml был создан
    app.kill()  # Закрываем приложение после выполнения теста


def test_import_packetepd():
    """
    Тест проверяет функциональность импорта файла PacketEPD.
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения
    window.menu_select("Import->PacketEPD")  # Открываем пункт меню «Import -> PacketEPD»
    file_dialog = app.window(title_re=".*Open.*")  # Находим диалоговое окно "Открыть файл"
    file_dialog.Edit.set_text(os.path.join(BASE_DIR, 'PacketEPD.xml'))  # Вводим путь к файлу PacketEPD.xml
    file_dialog.Open.click()  # Нажимаем кнопку "Открыть"
    assert "Импорт успешен" in window.static_texts()  # Проверяем, что отображается сообщение об успешном импорте
    # Здесь можно добавить дополнительные проверки содержимого таблицы
    app.kill()  # Закрываем приложение после выполнения теста


def test_export_to_csv():
    """
    Тест проверяет функциональность экспорта данных в CSV.
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения
    window.menu_select("Export->to CSV")  # Открываем пункт меню «Export -> to CSV»
    assert os.path.exists(os.path.join(BASE_DIR, 'ED101.csv'))  # Проверяем, что файл ED101.csv был создан
    app.kill()  # Закрываем приложение после выполнения теста


def test_export_to_json():
    """
    Тест проверяет функциональность экспорта данных в JSON.
    """
    app = start_application()  # Запускаем приложение
    window = app.window(title_re="Тестовое задание")  # Находим главное окно приложения
    window.menu_select("Export->to JSON")  # Открываем пункт меню «Export -> to JSON»
    assert os.path.exists(os.path.join(BASE_DIR, 'ED101.json'))  # Проверяем, что файл ED101.json был создан
    app.kill()  # Закрываем приложение после выполнения теста


if __name__ == '__main__':
    pytest.main()  # Запускаем pytest для выполнения тестов

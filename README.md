# Задание:
1.	В файле cities.xlsx таблица с названиями нескольких городов. Ваша программа должна с помощью веб-браузера найти информацию о погоде в них. Заполнить таблицу данными, сохранить под новым именем в формате cities_</date>_</time>.xlsx. Исходная таблица не должна быть изменена, форматирование должно быть сохранено.
cities.xlsx
2.	Программа должна открыть приложение Subline Text или аналогичное. Записать стихотворение Бармаглот, сохранить его на рабочий стол под любым названием. Отредактировать текущий файл, сохранить его на рабочий стол под другим названием. Закрыть Subline Text.

## Развертка проекта
1. Клонировать текущий репозиторий
2. Установить нужные зависимости при помощи команды `pip install -r requirements.txt`
3. Скачать и установить [chromedriver](https://chromedriver.chromium.org) нужной версии
4. В настройках компьютера разрешить программе управлять компьютером `Настройки -> Защита и безопастность -> Универсальный доступ`, снять ограничения, введя пароль и поставить галочку на против IDE, которой предоставляется доступ для запуска
5. В файле `base.py` в поле `PATH_CHROME_DRIVER` прописать путь, где располагается [chromedriver](https://chromedriver.chromium.org)
6. Проверить совпадает ли комбинация клавиш для открытия `Spotlight` в файле  `operations.py` в методе `open_program()`, в случае, если не совпадает изменить
7. Запустить программу

# import pyautogui
from operations import *


def task_1():
    cities = get_data_excel(base.PATH_FILE_EXCEL)
    data = get_data_from_browser(cities)
    save_data_excel(base.PATH_FILE_EXCEL, base.PATH_DESCTOP, data)


def task_2():
    open_program()
    add_text_in_file()
    save_as_file_poem()
    add_modify_in_file()
    save_as_file_poem()
    exit_program()


if __name__ == '__main__':
    logging.basicConfig(filename='app.log', filemode='w', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info('Start program')

    # Task 1
    task_1()

    # Task 2
    task_2()

    logging.info('End program')

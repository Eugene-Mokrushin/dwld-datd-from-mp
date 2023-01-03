import selenium.common.exceptions
from selenium import webdriver
import time
import openpyxl


# Перед тем как пользоваться программой, убедись что файл xlsx, драйвер и программа лежат в одной папке.
# Предварительно в xlsx файле разъедини столбцы, сделай отображение столбцов - буквами, сохрани файл и закрой(!).
# Обнови строчку 24.
# Не забыть обновить название файла на строках - 411, 457. (Строки могут менятся)
# В случае ошибки программы, перезапусти ее. Следи за консолью. Одна выдаст всю информацию.


def get_first_number(list):
    length_of_list = len(list)
    for val in range(1, length_of_list, 2):
        string_value = str(list[val])
        first_number = string_value[0]
        list[val] = int(first_number)
    return list


# Объясвление драйвера
browser = webdriver.Chrome('C:/Users/79150/Desktop/Selenium/chromedriver.exe')


# Shop4
def shop4_ozon():
    """Выгружает список товаров с магазина Ozon shop4"""

    try:
        # Осуществляет вход на сайт
        browser.get('https://seller.ozon.ru/signin')
        login_bar = browser.find_element_by_id('userName_1')
        login_bar.send_keys('*email*')
        password_bar = browser.find_element_by_id('password_2')
        password_bar.send_keys('*password*')
        log_in = browser.find_element_by_class_name('index_text_3dsim')
        log_in.click()
        time.sleep(2)

        # Переход на страницу заказов
        orders = browser.find_element_by_xpath('/html/body/div[1]/div/header/div[2]/div/div/div[5]/div/div[1]')
        orders.click()
        time.sleep(7)

        # Закрывает всплывающее окно
        close_ad = browser.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div/div[2]/div[2]/button/span')
        close_ad.click()
        time.sleep(5)

        # Переходит ко всем заказам (сборка)
        all_orders = browser.find_element_by_xpath(
            '/html/body/div[1]/div/main/div[1]/div[2]/div[3]/div/div[1]/div[1]/div[1]')
        all_orders.click()
        time.sleep(3)

        # Собирает список товаров ждущих сборки
        order = browser.find_elements_by_xpath("//div[@class='props']/span")
        sku_and_q = []
        for i in order:
            text = i.text
            sku_and_q.append(text)

        # Переходит к товаром ожидающих отправления
        waiting_for_shipment = browser.find_element_by_xpath(
            '/html/body/div[1]/div/main/div[1]/div[2]/div[2]/div/div/div/div/div[1]/div/div[3]/div[1]')
        waiting_for_shipment.click()
        time.sleep(3)
        waiting_for_shipment_all = browser.find_element_by_xpath(
            '/html/body/div[1]/div/main/div[1]/div[2]/div[3]/div/div[1]/div[1]/div[1]/div[1]')
        waiting_for_shipment_all.click()
        time.sleep(3)

        # Собирает список товаров ждущих отправления
        order_waiting_for_shipment = browser.find_elements_by_xpath("//div[@class='props']/span")
        sku_and_q_waiting = []
        for i in order_waiting_for_shipment:
            text = i.text
            sku_and_q_waiting.append(text)

        # Ожидает сборки
        x = get_first_number(sku_and_q)

        # Ожидает отгрузки
        y = get_first_number(sku_and_q_waiting)
        return x, y

    except selenium.common.exceptions.NoSuchElementException:
        return print('Shop4 не дал результата')


# Shop1
def shop0_ozon():
    """Не использовать без Shop0"""

    try:
        # Переходит на страницу входа
        browser.get('https://seller.ozon.ru/signin')

        # Осуществляет вход на сайт
        login_bar_1 = browser.find_element_by_id('userName_1')
        login_bar_1.send_keys('*email*')
        password_bar_1 = browser.find_element_by_id('password_2')
        password_bar_1.send_keys('*password*')
        log_in_1 = browser.find_element_by_class_name('index_text_3dsim')
        log_in_1.click()
        time.sleep(2)

        # Переходит в раздел заказов
        orders_1 = browser.find_element_by_xpath('/html/body/div[1]/div/header/div[2]/div/div/div[5]/div/div[1]')
        orders_1.click()
        time.sleep(5)

        # Использовать если возникла проблема с всплывающим окном
        """
        # time.sleep(5)
        # close_ad_1 = browser.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div/div[2]/div[2]/button/span')
        # close_ad_1.click()
        """

        # Переходит ко всем заказам
        """
        all_orders_1 = browser.find_element_by_xpath(
            '/html/body/div[1]/div/main/div[1]/div[2]/div[2]/div/div/div/div/div[1]/div/div[1]/div[1]')
        all_orders_1.click()
        time.sleep(3)
        """

        # Собирает список товаров ждущих сборки
        order_1 = browser.find_elements_by_xpath("//div[@class='props']/span")
        sku_and_q_1 = []
        for i in order_1:
            text = i.text
            sku_and_q_1.append(text)

        # Переходит к товаром ожидающих отправления
        waiting_for_shipment_1 = browser.find_element_by_xpath(
            '/html/body/div[1]/div/main/div[1]/div[2]/div[2]/div/div/div/div/div[1]/div/div[3]/div[1]')
        waiting_for_shipment_1.click()
        time.sleep(3)

        # Если появляется новая кнопка - все отправления
        """
        # waiting_for_shipment_all_1 = browser.find_element_by_xpath('/html/body/div[1]/div/main/div[1]/div[2]/div[3]/div/div[1]/div[1]/div[1]/div[1]')
        # waiting_for_shipment_all_1.click()
        # time.sleep(3)
        """

        # Собирает список товаров ждущих отправления
        order_waiting_for_shipment_1 = browser.find_elements_by_xpath("//div[@class='props']/span")
        sku_and_q_waiting_1 = []
        for i in order_waiting_for_shipment_1:
            text = i.text
            sku_and_q_waiting_1.append(text)

        # Ожидает сборки
        x = get_first_number(sku_and_q_1)

        # Ожидает отгрузки
        y = get_first_number(sku_and_q_waiting_1)

        return x, y

    except selenium.common.exceptions.NoSuchElementException:
        return print('*shop name* - Ozon не дал результата')


# Yandex 1
def shop2_yandex():
    """Выгружает список товаров с магазина Yandex Shop2"""

    try:
        # Заходит на страницу яндекс-магазина и последовательно вводит логин и пароль
        browser.get('https://passport.yandex.ru/auth/welcome?retpath=https://market.yandex.ru/partners')
        login_bar_yand = browser.find_element_by_id('passp-field-login')
        login_bar_yand.send_keys('*email*')
        enter_1 = browser.find_element_by_xpath(
            '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div[3]/div/div/div/div[1]/form/div[3]/button')
        enter_1.click()
        time.sleep(1)
        password_bar_yand = browser.find_element_by_id('passp-field-passwd')
        password_bar_yand.send_keys('*password*')
        time.sleep(1)
        enter_2 = browser.find_element_by_xpath(
            '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div[3]/div/div/div/form/div[3]/button')
        enter_2.click()
        time.sleep(1)
        time.sleep(1)

        # Переходит на страницу выбора магазина (!) Shop1 пока в коде не активен (!)
        browser.get('https://partner.market.yandex.ru/business/*shop1*/dashboard')
        time.sleep(1)

        # Переходит в Shop1
        shop1 = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div[2]/div/div/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[1]/td[1]/div/div/div[1]/a/span/span')
        shop1.click()
        time.sleep(3)

        # Переход в страницу заказов
        techno_orders = browser.find_element_by_xpath('//*[@id="app__sidebar"]/div/div[1]/div/nav/a[7]/span')
        techno_orders.click()

        # Выбирает перод заказов
        pick_a_date = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[2]/div/div/div[3]/div/div/div[2]/label/div/div/div/button/span/span')
        pick_a_date.click()
        time.sleep(0.5)
        current_d = browser.find_elements_by_css_selector('div.DayPicker-Day--today')
        current_d[1].click()
        future_d = browser.find_element_by_xpath(
            '/html/body/div[10]/div[2]/div/div/div/div/div[3]/div[3]/div[5]/div[3]')
        future_d.click()
        left_date = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[2]/div/div/div[1]/div/div/div/div/p/span')
        left_date.click()

        # Отоброжать 50 товаров на стрк
        fifty_orders = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[3]/div/div/div[3]/div/div/div[1]/div/ul/li[3]/a')
        fifty_orders.click()
        time.sleep(10)

        # Список статусов товара
        yand_list_of_orders_discription = []
        declined_orders = browser.find_elements_by_xpath("//div[@class='style-manager___39coh']/span/span")
        for i in declined_orders:
            text = i.text
            yand_list_of_orders_discription.append(text)

        # Список самих товаров
        yand_list_of_orders = []
        time.sleep(8)
        declined_orders = browser.find_elements_by_xpath("//div[@class='style-itemCaption___2OYTD']/span")
        for i in declined_orders:
            text = i.text
            yand_list_of_orders.append(text)
        final_yand_list_of_orders = []

        # Совмещает шт и SKU
        lenf = len(yand_list_of_orders)
        for i in range(0, lenf, 2):
            ful = yand_list_of_orders[i] + ' ' + yand_list_of_orders[i + 1]
            final_yand_list_of_orders.append(ful)
        denied_elements = []

        # Удаляет отмененные товары из списка SKU
        for i in range(len(yand_list_of_orders_discription)):
            if yand_list_of_orders_discription[i] == 'Отменён':
                denied_elements.append(i)
        for i in denied_elements:
            final_yand_list_of_orders[i] = None
        final_yand_list_of_orders.remove(None)
        r_item = None
        for item in final_yand_list_of_orders:
            if item == r_item:
                final_yand_list_of_orders.remove(r_item)

        # Проверка
        """
        print(final_yand_list_of_orders)
        print(len(yand_list_of_orders_discription), len(final_yand_list_of_orders))
        """

        return final_yand_list_of_orders

    except selenium.common.exceptions.NoSuchElementException:
        return print('Shop не дал результата')


# Yandex 2
# Shop1 Пока не используется
def Shop2_yandex():
    """Должен искользоваться с Shop1"""

    try:
        # Выход на стр выбор магазинов
        browser.get('https://partner.market.yandex.ru/business/*num*/dashboard')
        time.sleep(1)

        # Переход в Shop1
        shop1 = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div[2]/div/div/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[2]/td[1]/div/div/div[1]/a/span/span')
        shop1.click()
        time.sleep(3)

        # Переход в заказы
        techno_orders = browser.find_element_by_xpath('//*[@id="app__sidebar"]/div/div[1]/div/nav/a[7]/span')
        techno_orders.click()

        # Выбирает перод заказов
        pick_a_date = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[2]/div/div/div[3]/div/div/div[2]/label/div/div/div/button/span/span')
        pick_a_date.click()
        time.sleep(0.5)
        current_d = browser.find_elements_by_css_selector('div.DayPicker-Day--today')
        current_d[1].click()
        future_d = browser.find_element_by_xpath(
            '/html/body/div[10]/div[2]/div/div/div/div/div[3]/div[3]/div[5]/div[3]')
        future_d.click()
        left_date = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[2]/div/div/div[1]/div/div/div/div/p/span')
        left_date.click()

        # Отоброжать 50 товаров на стрк
        fifty_orders = browser.find_element_by_xpath(
            '//*[@id="app"]/div/div[1]/div/div/div/div/div[1]/div/div[3]/div/div/div[3]/div/div/div[1]/div/ul/li[3]/a')
        fifty_orders.click()
        time.sleep(10)

        # Список статусов товара
        yand_list_of_orders_discription_1 = []
        declined_orders = browser.find_elements_by_xpath("//div[@class='style-manager___39coh']/span/span")
        for i in declined_orders:
            text = i.text
            yand_list_of_orders_discription_1.append(text)

        # Список самих товаров
        yand_list_of_orders_1 = []
        time.sleep(8)
        declined_orders = browser.find_elements_by_xpath("//div[@class='style-itemCaption___2OYTD']/span")
        for i in declined_orders:
            text = i.text
            yand_list_of_orders_1.append(text)
        final_yand_list_of_orders_1 = []

        # Совмещает шт и SKU
        lenf = len(yand_list_of_orders_1)
        for i in range(0, lenf, 2):
            ful = yand_list_of_orders_1[i] + ' ' + yand_list_of_orders_1[i + 1]
            final_yand_list_of_orders_1.append(ful)
        denied_elements_1 = []

        # Удаляет отмененные товары из списка SKU
        for i in range(len(yand_list_of_orders_discription_1)):
            if yand_list_of_orders_discription_1[i] == 'Отменён':
                denied_elements_1.append(i)
        for i in denied_elements_1:
            final_yand_list_of_orders_1[i] = None
        final_yand_list_of_orders_1.remove(None)
        r_item = None
        for item in final_yand_list_of_orders_1:
            if item == r_item:
                final_yand_list_of_orders_1.remove(r_item)

        # Проверка
        """
        print(final_yand_list_of_orders_1)
        print(len(yand_list_of_orders_discription_1), len(final_yand_list_of_orders_1))
        """

        return final_yand_list_of_orders_1

    except selenium.common.exceptions.NoSuchElementException:
        return print('Shop - Yandex не дал результата')


all_shop4 = shop4_ozon()
print(f'Ozon 4 \n {all_shop4}')
print('\n')
shop0_ozon = shop0_ozon()
print(f'Ozon 0 \n {shop0_ozon}')
print('\n')
Shop2_yandex = Shop2_yandex()
print(f'Yandex 2 \n {Shop2_yandex}')
cont = input(
    'Продолжить работу программы?y/n (!) Если не из всех магазинов получена информация - перезапусти файл (!)\n')
if cont.lower() == 'y':

    '''Private DATA'''

else:
    print('Программа была завершена пользователем')

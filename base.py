from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import library


def initial(value_tb):                                               # Initial and open page 'url'
    path_driver = Service(r"C:\Users\13ins\Downloads\chromedriver_win32 (4)\chromedriver.exe")  # Driver Chrome path
    chrome_opt = Options()
    # chrome_opt.add_argument('--kiosk')              #Fullscreen
    browser = webdriver.Chrome(service=path_driver, options=chrome_opt)
    browser.maximize_window()
    try:
        browser.get('https://'+value_tb['value_1'])
    except Exception as exc:
        raise Exception('Failed switch to url') from exc
    return browser


def search_bar(browser):                               # Search Value_2 in Yandex
    path_search = '//*[@id="text"]'  # Search element Xpath
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, path_search)))
    try:
        elem_for_search = browser.find_element(By.XPATH, path_search)
    except Exception as exc:
        raise Exception('Search panel dont found') from exc
    elem_for_search.send_keys(use_value_tb['value_2'])
    elem_for_search.send_keys('\ue007')                                        # Push 'Enter'
    check_url_correct(browser)


def check_url_correct(browser):                          # url processing search , going to 1 non-marketing link
    url = library.parsing_url(browser)
    try:
        browser.get(url)
    except Exception as exc:
        raise Exception('no valid url') from exc
    test_title(browser)


def test_title(browser):                                # check equal title
    value_2 = use_value_tb['value_2']
    title_path = '/html/head/title'
    WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, title_path)))
    title_get = browser.find_element(By.XPATH, title_path).get_attribute('text')
    try:
        value_2 in title_get.split(' ')
    except Exception as exc:
        raise Exception('condition not met for title') from exc
    work_frame(browser)


def work_frame(browser):                                                  # transition on work page
    WebDriverWait(browser, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/div[1]/div/main/div/section[1]/a[5]/div'))
    )
    element_to_hover_over = browser.find_element(By.XPATH, '/html/body/div[1]/div/main/div/section[1]/a[5]/div')
    hover = library.ActionChains(browser).move_to_element(element_to_hover_over)
    hover.perform()
    path_vl_3 = '//*[@id="__next"]/div/header/nav[2]/div[2]/div[1]/div/a[2]'
    path_vl_4 = '//*[@id="__next"]/div/main/section[2]/div/ul/li[1]/a'
    browser.implicitly_wait(3)
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, path_vl_3)))
    browser.find_element(By.XPATH, path_vl_3).click()
    browser.find_element(By.XPATH, path_vl_4).click()


def iterator(browser, value_tb):                                        # making data for write
    value = value_tb['value_5']           # Data from table
    step = value_tb['value_6']
    end = value_tb['value_7']
    data_for_write = {}
    while value < end:
        data_for_write[value] = library.iterable_count(browser, value)
        value += step
    data_for_write[end] = library.iterable_count(browser, end)
    browser.save_screenshot('screenshot.png')      # Create and save screenshot
    library.write_data(data_for_write)             # Create and write in txt file data


def main():                                             # start func
    browser = initial(use_value_tb)
    search_bar(browser)

    iterator(browser, use_value_tb)
    browser.quit()                                 # Close window


first_table = 'value.xlsx'
list_name_for_first = 'Лист1'

First_tb = library.XlData(first_table, list_name_for_first)
use_value_tb = First_tb.dict_data
main()







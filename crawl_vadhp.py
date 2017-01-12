# -*- coding: utf-8 -*-
import time
import urllib
import xlsxwriter
import re
import datetime
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select

import sys
reload(sys)
sys.setdefaultencoding('utf8')


def crawl_document(url):

    time_interval = 1  # if network speed is slow, set more interval

    driver = webdriver.Chrome("C:/Users/Robin/Documents/Code/Python/package/chromedriver_win32/chromedriver.exe")

    driver.get(url)

    select = Select(driver.find_element_by_tag_name("select"))
    select.select_by_visible_text("Medicine")

    driver.find_element_by_name("txtBDate").send_keys("12-01-2016")    # set the beginning data
    driver.find_element_by_name("txtEDate").send_keys("12-15-2016")    # set the ending data

    time.sleep(time_interval)

    search = driver.find_element_by_name("send")
    search.submit()

    time.sleep(time_interval+1)  # search result page need more time to load, if needed, add more seconds

    wb_data = driver.execute_script("return document.documentElement.innerHTML")
    soup = BeautifulSoup(wb_data, 'lxml')

    now_handle = driver.current_window_handle  # the search page window

    view_buttons = driver.find_elements_by_id('submit1')  # find the 'view' button

    license_list = []
    with open('licensedata.txt', 'r') as license_data:
        for txt_data in license_data.readlines():
            license_number = txt_data.strip('\n')
            license_list.append(license_number)
            license_data.close()
    license_list = list(set(license_list))   # get the all license numbers already crawled

    row = 0
    col = 0
    workbook = xlsxwriter.Workbook('document_000.xlsx')  # create a excel file to save information
    worksheet = workbook.add_worksheet()

    for i in range(0, len(view_buttons)):

        j = 2 * i

        license_number = soup.find_all('td', class_='xl24')[j].getText()

        if license_number in license_list:
            print 'Already download'
            continue
        else:
            print 'New record'
            license_list.append(license_number)

        view_button = view_buttons[i]

        view_button.click()

        time.sleep(time_interval)

        for handle in driver.window_handles:         # switch to new window
            driver.switch_to.window(handle)

        driver.find_element_by_name("send").click()

        time.sleep(time_interval)

        wb_data_download = driver.execute_script("return document.documentElement.innerHTML")
        soup_download = BeautifulSoup(wb_data_download, 'lxml')

        data_urls = soup_download.select('font > a')

        for data_url in data_urls:
            file_url = data_url.get('href')
            file_name = file_url.split('/')[-1]

            if file_name == "readstep.html":
                pass
            else:
                file_licnum = file_name[0:10]
                file_type = ''.join(re.findall(r'[a-zA-Z]', file_name.split('.')[0]))
                file_date_data = ''.join(reversed(file_name[-5:-13:-1]))
                try:

                    file_date = datetime.datetime.strptime(file_date_data, '%m%d%Y').strftime('%m/%d/%Y')

                except Exception as e:
                    print e

                print file_licnum, file_type, file_date
                print file_url, file_name

                save_path = 'C:/Users/Robin/Documents/Code/Project/crawl_vadhp/download/' + file_name
                # to set the download folder
                urllib.urlretrieve(file_url, save_path)

                time.sleep(time_interval)

                worksheet.write(row, col, file_licnum)
                worksheet.write(row, col + 1, file_type)
                worksheet.write(row, col + 2, file_date)
                worksheet.write(row, col + 3, file_name)

                row += 1

        driver.close()

        driver.switch_to.window(now_handle)   # switch to search page



    with open('licensedata.txt', 'w') as license_data:
        for license_number in license_list:
            license_data.write(license_number+'\n')
        license_data.close()


    driver.quit()
    workbook.close()

if __name__ == "__main__":

    print "Begin to download"
    crawl_document('https://www.dhp.virginia.gov/enforcement/cdecision/cd_advsearch.asp')

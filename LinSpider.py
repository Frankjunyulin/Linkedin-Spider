#Written by Frank Lin in 2017.
from selenium import webdriver
from bs4 import BeautifulSoup
import json
import csv
import time
from random import randint
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from getDataFunc import *





def main():

    username = input("Linkedin Username: ")
    password = input("Linkedin Password: ")
    path = input("input file: ")
    fileName = input("output file: ")


    if os.path.isfile(fileName):
        print('old file')
    else:
        print('new file')
        write_file('example.xlsx')

    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome('chromedriver', chrome_options=chrome_options)
    #driver = webdriver.Firefox()


    driver.get('https://www.linkedin.com/')
    driver.set_page_load_timeout(100)
    driver.find_element_by_name("session_key").send_keys(username)
    driver.find_element_by_name("session_password").send_keys(password)
    driver.find_element_by_id("login-submit").click()


    search_List = open_file(path)
    keyWord = ''
    for search_pro in search_List:

        search_id = search_pro[0]
        search_first_name = search_pro[1]
        search_last_name = search_pro[2]
        keyWord = search_first_name + ' ' + search_last_name
        keyWord = keyWord.replace(" ", "%20")

        searchURL = 'https://www.linkedin.com/search/results/people/?facetGeoRegion=%5B%22us%3A0%22%5D&facetIndustry=%5B%22106%22%2C%2243%22%2C%2241%22%2C%2242%22%2C%2246%22%2C%2245%22%2C%22129%22%5D' '&keywords='+ keyWord + '&origin=FACETED_SEARCH'

        links = []
        page_number = 1

        while True:
            pageURL = searchURL + '&page=' + str(page_number)
            driver.get(pageURL)

            #test whether result exists
            if driver.find_elements_by_css_selector(".search-no-results"):
                print("no results")
                break


            driver.implicitly_wait(10)

            i = 0
            while i<1000:
                scroll = driver.find_element_by_tag_name('body').send_keys(Keys.END)
                i+=1

            sleep(5)

            element_present = EC.presence_of_element_located((By.CSS_SELECTOR, '.search-result__info'))
            WebDriverWait(driver, 30).until(element_present)

            results = driver.find_elements_by_css_selector(".search-result__info")

            for r in results:
                this_link = r.find_element_by_css_selector('a').get_attribute('href')
                print(this_link)
                if '/in/' in this_link:
                    links.append(this_link)
                    print("These are links:")
                    print(this_link)

            page_number += 1

        for temp_link in links:
            anchor = temp_link
            getData(anchor, driver, search_first_name, search_last_name, search_id)




if __name__ == '__main__':
    main()

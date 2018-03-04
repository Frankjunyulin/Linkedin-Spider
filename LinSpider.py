#Written by Junyu Lin in 2017.


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
from fileFunc import *


def getData(link, driver, search_first_name, search_last_name, search_id):

    driver.get(link)
    driver.set_page_load_timeout(30)
    html = driver.page_source
    driver.implicitly_wait(3)
    driver.execute_script("window.scrollTo(0, (document.body.scrollHeight)*3/5);")


    #scroll down
    try:
        scroll_element1 = driver.find_element_by_css_selector(".pv-top-card-section__school")
        scroll_element1.location_once_scrolled_into_view
    except:
        pass

    try:
        scroll_element2 = driver.find_element_by_css_selector(".education-section")
        scroll_element2.location_once_scrolled_into_view
    except:
        passpass

    try:
        scroll_element3 = driver.find_element_by_css_selector(".volunteering-section")
        scroll_element3.location_once_scrolled_into_view
    except:
        passpass

    try:
        scroll_element4 = driver.find_element_by_css_selector(".pv-featured-skills-section")
        scroll_element4.location_once_scrolled_into_view
    except:
        pass

    try:
        scroll_element5 = driver.find_element_by_css_selector(".pv-accomplishments-section")
        scroll_element5.location_once_scrolled_into_view
    except:
        pass


    try:
        scroll_element6 = driver.find_element_by_css_selector(".pv-interests-section")
        scroll_element6.location_once_scrolled_into_view
    except:
        pass


    sleep(5)

#####################################################
    soup = BeautifulSoup(html, "html.parser")
    json_alldata = soup.find_all("code")
    for x in json_alldata:
        try:
            json_data = json.loads(x.text)
        except ValueError:
            pass
        if "data" in json_data:
            if "$type" in json_data["data"]:
                if json_data["data"]["$type"] == "com.linkedin.voyager.identity.profile.ProfileView":
                    break

    name = headline = industry = location = summary = vanity_url = website = email = phone = "None"
    first_name= last_name= pro_title= pro_company= location = pro_school = follow_num = ''
    for x in json_data["included"]:
        if x["$type"] == "com.linkedin.voyager.identity.profile.Profile":
            try :
                name = x["firstName"] + " " + x["lastName"]
                first_name = x["firstName"]
                last_name = x["lastName"]
            except:
                pass

            try :
                headline = x["headline"]
                temp = headline.split(" at ")
                pro_title = temp[0]
                pro_company = temp[1]
            except:
                pass

            try :
                industry = x["industryName"]
            except:
                pass

            try :
                location = x["locationName"]
            except:
                pass

            try :
                summary = x["summary"]
            except:
                pass

    new_html = driver.page_source
    soup = BeautifulSoup(new_html, "html.parser")


    #Get other summary
    try:
        follower_num_getter = driver.find_element_by_css_selector(".pv-top-card-section__connections--with-separator")
        follow_num = follower_num_getter.find_element_by_xpath(".//span[not(contains(@class, 'svg-icon-wrap'))]").text
    except:
        pass

    try:
        pro_school = driver.find_element_by_css_selector(".pv-top-card-section__school").text
        if not pro_school:
            temp_list = pro_school.split('\"')
            pro_school = temp_list[1]
    except:
        pass

    try:
        summary_List = [search_id, search_first_name, search_last_name, '', first_name, last_name, pro_title , pro_company, '', '', location, pro_school, follow_num, '', link];
        edit_file(file_path, 0, summary_List)
    except:
        pass
    ###Summary var: first_name, last_name, pro_title, pro_company, location, follow_num, link


    #Extract experience
    exp_title = exp_company = exp_beginMon = exp_beginYear = exp_endMon = exp_endYear = exp_state = ''

    #More experience Button
    try:
        expButtonDivList = driver.find_elements_by_css_selector(".pv-profile-section__actions-inline")
    except:
        pass

    for expButtonDiv in expButtonDivList:
        try:
            expButton = expButtonDiv.find_element_by_xpath(".//button")
            while expButton.get_attribute('aria-expanded')  != 'true':
                expButton.send_keys(u'\ue007');
                expButton = expButtonDiv.find_element_by_xpath(".//button")
        except:
            pass




    exp_title = exp_company = exp_beginMon = exp_beginYear = exp_endMon = exp_endYear = exp_state = ''
    expList = []
    try:
        expList = driver.find_elements_by_css_selector(".pv-profile-section__card-item.pv-position-entity.ember-view")
    except:
        pass

    for exp in expList:
        try:
            exp_title = exp.find_element_by_xpath(".//a/div/h3").text
        except:
            pass
        try:
            exp_company = exp.find_element_by_xpath(".//a/div/h4[contains(@class, 'Sans-17px-black-85%')]/span[contains(@class, 'pv-entity__secondary-title')]").text
        except:
            pass

        try:
            exp_period = exp.find_element_by_xpath(".//a/div/h4[contains(@class, 'pv-entity__date-range')]/span[not(contains(@class, 'visually-hidden'))]").text
            exp_period = exp_period.replace('Present','2020')
            if ' – ' in exp_period:
                temp = exp_period.split(' – ')
                if ' ' in temp[0]:
                    temp1 = temp[0].split(' ')
                    exp_beginMon = temp1[0]
                    exp_beginYear = temp1[1]
                else:
                    exp_beginYear = temp[0]
                if ' ' in temp[1]:
                    temp2 = temp[1].split(' ')
                    exp_endMon = temp2[0]
                    exp_endYear = temp2[1]
                else:
                    exp_endYear = temp[1]
        except:
            pass



        try:
            exp_city = exp.find_element_by_xpath(".//a/div/h3").text
            exp_state = exp.find_element_by_xpath(".//a/div/h4[contains(@class, 'pv-entity__location')]/span[not(contains(@class, 'visually-hidden'))]").text
            exp_country = exp.find_element_by_xpath(".//a/div/h3").text
        except:
            pass

        try:
            exp_List = [search_id,'',first_name,last_name,exp_company,exp_title,exp_beginYear,exp_beginMon,exp_endYear,exp_endMon,exp_state,12,13]
            edit_file(file_path, 1, exp_List)
        except:
            pass

    #########var: exp_title, exp_company, exp_beginMon, exp_beginYear, exp_endMon, exp_endYear, exp_state

    #Extract Education
    edu = edu_beg_year =edu_beg_month = edu_end_year = edu_end_month = edu_degree = edu_major = ''
    edList = []
    try:
        edu_Section = driver.find_element_by_css_selector(".pv-profile-section.education-section.ember-view")
        edList = edu_Section.find_elements_by_xpath(".//ul/li")
    except:
        pass


    for ed in edList:
        try:
            edu = ed.find_element_by_xpath(".//div/a/div[contains(@class, 'pv-entity__summary-info')]/div/h3").text
        except:
            pass

        try:
            edu_time = ed.find_elements_by_xpath(".//div/a/div[contains(@class, 'pv-entity__summary-info')]/p/span[not(contains(@class, 'visually-hidden'))]/time")
            edu_beg_year= edu_time[0].text
            edu_end_year= edu_time[1].text
        except:
            pass

        try:
            edu_degree = ed.find_element_by_xpath(".//div/a/div[contains(@class, 'pv-entity__summary-info')]/div/p[contains(@class, 'pv-entity__degree-name')]/span[not(contains(@class, 'visually-hidden'))]").text
        except:
            pass

        try:
            edu_major = ed.find_element_by_xpath(".//div/a/div[contains(@class, 'pv-entity__summary-info')]/div/p[contains(@class, 'pv-entity__fos')]/span[not(contains(@class, 'visually-hidden'))]").text
        except:
            pass

        try:
            edu_List =  [search_id, '', first_name, last_name, edu, edu_degree, edu_beg_year , edu_beg_month, edu_end_year,edu_end_month,edu_major, '', '', '', ''];
            edit_file(file_path, 2, edu_List)
        except:
            pass

    ## Var: edu, edu_beg_year, edu_end_year,edu_degree, edu_major




    #####################################################
    #Extract Skills
    try:
        driver.find_element_by_css_selector(".pv-skills-section__additional-skills").click()
    except:
        pass


    skillName = ''
    skill_url_list = []
    try:
        skillList = driver.find_elements_by_css_selector(".featured-skill-entity-wrapper")
    except:
        pass
    for skill in skillList:
        try:
            skillName = skill.find_element_by_xpath(".//div/span[contains(@class, 'pv-skill-entity__skill-name')]")
        except:
            pass
        try:
            skill_url = skill.get_attribute("href")
            skill_url_list.append(skill_url)
        except:
            pass


    #Extract Accomplishments
    try:
        accomplishments = driver.find_element_by_css_selector(".pv-accomplishments-section")
    except:
        pass

    #certifications:
    cer_name = cer_time = ''
    certificationList = []
    try:
        certifications = accomplishments.find_element_by_css_selector(".certifications")
    except:
        pass
    try:
        certificationsButton = certifications.find_element_by_xpath(".//div/button")
        certificationsButton.click()
    except:
        pass
    try:
        certificationList = certifications.find_elements_by_xpath(".//div/div/ul/li")
    except:
        pass


    for certification in certificationList:
        try:
            cer_name = certification.find_element_by_xpath(".//h4").text
            cer_name_rubb = certification.find_element_by_xpath(".//h4/span").text
            cer_name = cer_name.replace(cer_name_rubb,'')
            cer_name = cer_name.strip('\n')
        except:
            pass

        try:
            cer_time = certification.find_element_by_xpath(".//p/span").text
            cer_time_rubb = certification.find_element_by_xpath(".//p/span/span").text
            cer_time = cer_time.replace(cer_time_rubb,'')
            cer_time = cer_time.strip('\n')
        except:
            pass
        try:
            accomp_List =  [pro_id, '', pro_first, pro_last, cer_name, cer_time, '' , '', '', '', '']
            edit_file(file_path, 6, accomp_List)
        except:
            pass

        #var: cer_name, cer_time

    #languages
    lan_name=''
    languagesList = []
    try:
        languages = accomplishments.find_element_by_css_selector(".pv-accomplishments-block.languages")
    except:
        pass
    try:
        languagesButton = languages.find_element_by_xpath(".//div/button")
        languagesButton.click()
    except:
        pass
    try:
        languagesList = languages.find_elements_by_xpath(".//div/div/ul/li")
    except:
        pass


    for language in languagesList:
        try:
            lan_name = language.find_element_by_xpath(".//h4").text
            lan_name_rubb = language.find_element_by_xpath(".//h4/span").text
            lan_name = lan_name.replace(lan_name_rubb,'')
            lan_name = lan_name.strip('\n')
        except:
            pass

        try:
            accomp_List =  [pro_id, '', pro_first, pro_last, '', '', '' , '', '', '', lan_name]
            edit_file(file_path, 6, accomp_List)
        except:
            pass
        #var: lan_name

    #awards
    award_name = award_time = award_institution = ''
    awardsList = []
    try:
        awards = accomplishments.find_element_by_css_selector(".pv-accomplishments-block.honors")
        awardsButton = awards.find_element_by_xpath(".//div/button")
        awardsButton.click()
    except:

        pass
    try:
        awardsList = awards.find_elements_by_xpath(".//div/div/ul/li")
    except:
        pass

    try:
        awaButtonDiv = accomplishments.find_element_by_css_selector(".pv-profile-section__actions-inline")
        awaButton = awaButtonDiv.find_element_by_xpath(".//button")
        while awaButton.get_attribute('aria-expanded')  != 'true':
            awaButton.send_keys(u'\ue007');
            awaButton = awaButtonDiv.find_element_by_xpath(".//button")
    except:
        pass

    try:
        awardsList = awards.find_elements_by_xpath(".//div/div/ul/li")
    except:
        pass

    for award in awardsList:
        try:
            award_name = award.find_element_by_xpath(".//h4").text
            award_name_rubb = award.find_element_by_xpath(".//h4/span").text
            award_name = award_name.replace(award_name_rubb,'')
            award_name = award_name.strip('\n')
        except:
            pass

        try:
            award_time = award.find_element_by_xpath(".//p/span[contains(@class, 'pv-accomplishment-entity__date')]").text
            award_time_rubb = award.find_element_by_xpath(".//p/span[contains(@class, 'pv-accomplishment-entity__date')]/span").text
            award_time = award_time.replace(award_time_rubb,'')
            award_time = award_time.strip('\n')
        except:
            pass

        try:
            award_institution = award.find_element_by_xpath(".//p/span[contains(@class, 'pv-accomplishment-entity__issuer')]").text
            award_institution_rubb = award.find_element_by_xpath(".//p/span[contains(@class, 'pv-accomplishment-entity__issuer')]/span").text
            award_institution = award_institution.replace(award_institution_rubb,'')
            award_institution = award_institution.strip('\n')
        except:
            pass

        try:
            accomp_List =  [pro_id, ' ', pro_first, pro_last, ' ', ' ', ' ' , award_name, award_time, award_institution, '']
            edit_file(file_path, 6, accomp_List)
        except:
            pass
        #var: award_name, award_time, award_institution

    #Extract People also viewed
    viewd_name = viewed_first_name = viewed_last_name = viewed_title = viewed_company = ''
    viewedList = []
    try:
        viewedList = driver.find_elements_by_css_selector(".pv-browsemap-section__member-container")
    except:
        pass


    for viewed in viewedList:
        try:
            viewed_name = viewed.find_element_by_xpath(".//a/div/h3/span[contains(@class, 'name-and-icon')]/span").text
            viewed_name_list = viewed_name.split(' ')
            viewed_first_name = viewed_name_list[0]
            viewed_last_name = viewed_name_list[-1]
        except:
            pass

        try:
            viewed_info = viewed.find_element_by_xpath(".//a/div/p").text
            viewed_title = viewed_info
        except:
            pass

        try:
            viewed_List =  [pro_id,'',pro_first,pro_last,viewed_first_name,viewed_last_name,'',viewed_title,viewed_company,'']
            edit_file(file_path, 7, viewed_List)
        except:
            pass
        #viewd_name, viewed_title


    for skill_url in skill_url_list:
        driver.get(skill_url)

        try:
            skillName = driver.find_element_by_css_selector(".pv-profile-detail__title").text
            temp0 = skillName.split("(")
            skillName = temp0[0]
        except:
            pass

        try:
            temp1 = temp0[1].split(")")
            endor_num = int(temp1[0])
            if endor_num > 9:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except:
            pass


        try:
            endor_list = driver.find_elements_by_css_selector(".pv-endorsement-entity.pv3.ember-view")
        except:
            pass

        for endor in endor_list:
            try:
                endor_name = endor.find_element_by_xpath(".//a/div/div[contains(@class, 'pv-endorsement-entity__name')]/span[contains(@class, 'pv-endorsement-entity__name--has-hover')]").text
                temp = endor_name.split(' ')
                endor_first_name = temp[0]
                endor_last_name = temp[-1]
            except:
                pass
            try:
                endor_title = endor.find_element_by_xpath(".//a/div/div[contains(@class, 'pv-endorsement-entity__headline')]").text
            except:
                pass

            try:
                skill_List = [pro_id, '', pro_first, pro_last, skillName, endor_first_name, endor_last_name , '', endor_title, '', '']
                edit_file(file_path, 3, skill_List)
            except:
                pass


    #GET FOLLOWING Influencers
    following_link = link + "detail/interests/influencers/"
    driver.get(following_link)

    inf_first_name = inf_last_name = inf_title = inf_company = inf_url = inf_num_followers = ''
    influList = []
    try:
        influList = driver.find_elements_by_css_selector(".pv-interest-entity-link")
    except:
        pass

    for influ in influList:
        try:
            inf_url = influ.get_attribute("href")
        except:
            pass

        try:
            inf_name = influ.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/h3/span").text
            inf_name_list = inf_name.split(' ')
            inf_first_name = inf_name_list[0]
            inf_last_name = inf_name_list[-1]
        except:
            pass

        try:
            inf_TilCom = influ.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/p[contains(@class, 'pv-entity__occupation')]").text
            inf_title = inf_TilCom
        except:
            pass

        try:
            inf_num_followers = influ.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/p[contains(@class, 'pv-entity__follower-count')]").text
        except:
            pass

        try:
            influ_List =  [pro_id,'',pro_first,pro_last,inf_first_name, inf_last_name, inf_title, inf_company, '', '', inf_url,inf_num_followers,'']
            edit_file(file_path, 4, influ_List)
        except:
            pass



    #get following companies
    following_link = link + "detail/interests/companies/"
    driver.get(following_link)

    follCom_name = ''
    try:
        follComList = driver.find_elements_by_css_selector(".pv-interest-entity-link")
    except:
        pass

    for follCom in follComList:
        try:
            follCom_name = follCom.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/h3/span").text
        except:
            pass
        try:
            com_num_followers = follCom.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/p[contains(@class, 'pv-entity__follower-count')]").text
        except:
            pass
        try:
            company_List =  [pro_id,'',pro_first,pro_last,follCom_name,'','']
            edit_file(file_path, 5, company_List)
        except:
            pass


    #GET FOLLOWING GROUPS
    following_link = link + "detail/interests/groups/"
    driver.get(following_link)
    follGroup_name = ''
    try:
        follGroupList = driver.find_elements_by_css_selector(".pv-interest-entity-link")
    except:
        pass

    for follGroup in follGroupList:
        try:
            follGroup_name = follGroup.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/h3/span").text
        except:
            pass
        try:
            group_num_followers = follGroup.find_element_by_xpath(".//div[contains(@class, 'pv-entity__summary-info')]/p[contains(@class, 'pv-entity__follower-count')]").text
        except:
            pass
        try:
            group_List =  [pro_id,'',pro_first,pro_last,'','',follGroup_name]
            edit_file(file_path, 5, group_List)
        except:
            pass




    #get following school
    following_link = link + "detail/interests/schools/"
    driver.get(following_link)


    school_name = ''

    #schoolList = driver.find_elements_by_css_selector(".pv-entity__summary-title-text")
    try:
        schoolList = driver.find_elements_by_css_selector(".interest-content")
    except:
        pass

    for school in schoolList:
        try:
            school_name = school.find_element_by_xpath(".//a/div[contains(@class, 'pv-entity__summary-info')]/h3/span").text
        except:
            pass
        try:
            school_foll = school.find_element_by_xpath(".//a/div[contains(@class, 'pv-entity__summary-info')]/p[contains(@class, 'pv-entity__follower-count')]").text
        except:
            pass
        try:
            school_List =  [pro_id,'',pro_first,pro_last,'',school_name,'']
            edit_file(file_path, 5, school_List)
        except:
            pass



    with open('data.csv', 'ab') as f:
        writer = csv.writer(f)
        #writer.writerow([name.encode("ascii", "ignore"), headline.encode("ascii", "ignore"), industry.encode("ascii", "ignore"), location.encode("ascii", "ignore"), summary.encode("ascii", "ignore"), vanity_url, website, email, phone])



############################


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

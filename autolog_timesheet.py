# -*- coding: utf-8 -*-
"""
Created on Mon Jul 10 13:17:31 2017
@author: kchen
"""

from bs4 import BeautifulSoup
import io
from selenium import webdriver


start_date = '2017-08-06'
end_date = '2017-08-17'

def parse_speakhi():
    # create a new Chrome session
    browser = webdriver.Chrome()
    
    browser.get("http://speakhi.com/web/teacher/teacher_login.html")
    username = browser.find_element_by_id("login_teacher_account")
    password = browser.find_element_by_id("login_teacher_pwd")
    username.send_keys('nicholas')
    password.send_keys('111111')
    browser.find_element_by_class_name("hwj-interaction-lg-button").click()
    browser.refresh
    time.sleep(10)
    html_source = browser.page_source
    courses = browser.find_element_by_class_name("course-content").text.split('\n')

def parse_html(html_file):
    parsed_html = BeautifulSoup(html_file, 'lxml')
    teacher = parsed_html.body.find('h3', attrs={'id':'teachername'}).text
    timeMonth = parsed_html.body.find_all('div', attrs={'class':'timeMonth'})
    timeHour = parsed_html.body.find_all('div', attrs={'class':'timeHour'})
    roomAstudent = parsed_html.body.find_all('ul', attrs={'class':'roomAstudent'})
    n = len(timeMonth)
    csv_file = io.open(teacher + '_new_courses.csv','w',encoding = 'utf-8-sig')
    for i in range(0, n):
        csv_file.write(timeMonth[i].text)
        csv_file.write(',')
        csv_file.write(',')
        csv_file.write(timeHour[i].text)
        csv_file.write(',')
        csv_file.write(roomAstudent[i].findAll('li')[1].text[0])
        csv_file.write('\n')
    
    csv_file.close()

def web_scrap(name, pin, start_date, end_date, browser):
    # navigate to the application home page
    browser.get("http://t.webi.com.cn/logon")
    username = browser.find_element_by_id("username")
    password = browser.find_element_by_id("password")

    username.send_keys(name)
    password.send_keys(pin)
    
    browser.find_element_by_class_name("loginButton").click()
    browser.get("http://t.webi.com.cn/course/list/0")
    
    
    browser.get('http://t.webi.com.cn/course/list/3?startDate=' + start_date + '&endDate=' + end_date)
    html_source = browser.page_source
    parse_html(html_source)
    #logout
    browser.get("http://t.webi.com.cn/home/logout")

if __name__ == "__main__":
    # create a new Chrome session
    browser = webdriver.Chrome()
    
    with open('login_credentials.txt') as f:
        for line in f:
            [teacher, name, pin] = line.split()
            web_scrap(name, pin, start_date, end_date, browser)
            
    browser.close()
    

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 22:11:02 2017

@author: kuilinchen
"""
import time
from bs4 import BeautifulSoup
import io
from selenium import webdriver

login_url = "https://collab.torontomls.net/signin"

# create a new Chrome session
browser = webdriver.Chrome("/usr/local/Cellar/chromedriver/2.34/bin/chromedriver") 
browser.get(login_url)
time.sleep(10)
username = browser.find_element_by_id("username")
password = browser.find_element_by_id("password")
username.send_keys("alexalex222")
time.sleep(10)
password.send_keys("1123581321")
time.sleep(10)
browser.find_element_by_class_name("btn-primary").click()
browser.refresh
time.sleep(10)
browser.find_element_by_class_name("icon-lists").click()

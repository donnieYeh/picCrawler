from asyncore import loop
from concurrent.futures import process
from email import header
import os
from pickle import TRUE
import time
from tokenize import group
from win32com.client.gencache import EnsureDispatch as Dispatch
import re
from urllib import parse as urlParse
from urllib import request as urlRequest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from bs4 import BeautifulSoup as bs

campaignSet = set()
def getUnreadMails():
    outlook = Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    Accounts = mapi.Folders  # 根级目录（邮箱名称，包括Outlook读取的存档名称）
    for Account_Name in Accounts:
        if Account_Name.Name != "xr08255920@gmail.com":
            continue
        print(' >> 正在查询的帐户名称：', Account_Name.Name, '\n')
        L1Foloders = Account_Name.Folders
        for L1 in L1Foloders:
            if L1.Name == "私人邮件":
                L2Folders = L1.Folders
                for L2 in L2Folders:
                    if L2.Name == "pinterest":
                        mails = L2.Items
                        break
                break
    print(len(mails))
    mails.Sort("ReceivedTime", True)
    return mails

def traverseMails(mails):
    mail = mails.GetFirst()
    while(mail != None):
        processMail(mail)
        mail = mails.GetNext()

def processMail(mail):
     mailContent = mail.Body
     urls = getURLsFromContent(mailContent)
     for url  in urls:
         processUrl(url)

def processUrl(url):
    if(not re.search("utm_campaign",url)):
        return
    query = urlParse.urlparse(url).query
    result = None
    count = 0
    while(not result and count < 5):
        query = urlParse.unquote_plus(query)
        result = re.search("utm_campaign=(.+?)&",query)
        count+=1

    if result:
        campaignSet.add(result.group(1))

def getURLsFromContent(mailContent):
    result = re.findall("<https://.*?>", mailContent)
    urlset = set()
    for i in result:
        urlset.add(i[1:-1])
    return urlset

mails = getUnreadMails()
traverseMails(mails)
print(campaignSet)
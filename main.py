from curses import raw
from email import header
import os
import time
from win32com.client.gencache import EnsureDispatch as Dispatch
import re
from urllib import parse as urlParse
from urllib import request as urlRequest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from bs4 import BeautifulSoup as bs


maxMailCount = 5
header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'}
maxCrawlSize = 30
filterstr = "[UnRead] = true"  


def main(targetPath):
    # 获取邮件未读热门pinterest内容列表
    mailContentList = getUnreadMails()
    # 解析邮件内容，获取所有分类链接
    categroyList = getCategories(mailContentList)
    urls = convertToRawList(categroyList)
    # 模拟浏览器打开网页，模拟滚动，获取所有图片地址，/pin/。。
    imgUrls = crawlImgs(urls)
    # 访问图片地址，获取原图链接
    rawImgUrls = toOriginalUrl(imgUrls)
    # 下载原图链接到目标文件夹
    download(rawImgUrls, targetPath)
    print("总共处理：",len(rawImgUrls))


def getUnreadMails():
    mailList = []
    outlook = Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    Accounts = mapi.Folders  # 根级目录（邮箱名称，包括Outlook读取的存档名称）
    for Account_Name in Accounts:
        if Account_Name.Name != "xr08255920@gmail.com":
            continue
        print(' >> 正在查询的帐户名称：', Account_Name.Name, '\n')
        L1Foloders = Account_Name.Folders
        for L1 in L1Foloders:
            if L1.Name != "私人邮件":
                continue
            L2Folders = L1.Folders
            for L2 in L2Folders:
                if L2.Name != "pinterest":
                    continue
                mails = L2.Items
                print(len(mails))
                mails.Sort("ReceivedTime", True)
                count = 0
                while count < maxMailCount:
                    mail = mails.FindNext()
                    if mail == None:
                        mail = mails.Find(filterstr)
                    if "_MailItem" not in str(type(mail)):
                        print("没有邮件可读")
                        break
                    mailList.append(mail)
                    count +=1
                # print(mails[1].Body)
                # mail.UnRead = False
    return mailList

# 处理两件事：1、只拉recommend的；2、只拉带utm_content的链接


def getCategories(mailList):
    categoryList = set()
    for mail in mailList:
        currentSet = set()
        print("正在处理邮件：", mail.Subject)
        mail.UnRead = False
        if "recommend" not in mail.SenderEmailAddress:
            print("该邮件不是推荐图片，跳过：", mail.Subject)
            continue
        mailCotent = mail.Body
        urls = getURLsFromContent(mailCotent)
        for url in urls:
            if "utm_content" in url:
                currentSet.add(url)
        print("该件提取url数：", len(currentSet))
        categoryList.update(currentSet)
    return categoryList


def getURLsFromContent(mailContent):
    result = re.findall("<https://.*?>", mailContent)
    urlset = set()
    for i in result:
        urlset.add(i[1:-1])
    return urlset


def convertToRawList(categoryList):
    rawUrlList = []
    for url in categoryList:
        url = urlParse.unquote(urlParse.unquote(url))
        paramHeader = "&next="
        index = url.index(paramHeader)
        url = url[index+len(paramHeader):]
        url = urlParse.urljoin("https://www.pinterest.com", url)
        rawUrlList.append(url)
    return rawUrlList


def crawlImgs(categoryList):
    imgs = set()
    browser = webdriver.Chrome()
    high = browser.execute_script(
        "return document.scrollingElement.clientHeight")
    high = str(high)
    print("浏览器高度：", high)
    for img in categoryList:
        if urlParse.urlparse(img).path.startswith("/pin/"):
            print("加入pin图:", img)
            imgs.add(img)
            continue
        pins = set()
        print("开始抓取图板：", img)
        try:
            browser.get(img)
        except:
            print("该连接打开:", img)
            continue
        browser.implicitly_wait(5)
        while(len(pins) < maxCrawlSize or browser.execute_script("document.scrollingElement.scrollTop+document.scrollingElement.clientHeight == document.scrollingElement.scrollHeight") == False):
            if len(pins) != 0:
                browser.execute_script("window.scrollBy(0,"+high+")")
                browser.implicitly_wait(3)
            elements = browser.find_elements(
                By.XPATH, '//*[@data-test-id="feed"]//a[contains(@href,"/pin/")]')
            for i in elements:
                try:
                    pins.add(i.get_attribute("href"))
                except StaleElementReferenceException:
                    continue
            print("\r已抓取数量：", len(pins), end="")
        print("")
        imgs.update(pins)
    browser.close()
    return imgs


def toOriginalUrl(imgUrls):
    rawImgs = set()
    for url in imgUrls:
        print("开始处理图片：",url)
        try:
            request = urlRequest.Request(url, headers=header)
            resp = urlRequest.urlopen(request)
            data = resp.read()
            soup = bs(data, 'html.parser')
            links = soup.find_all("link")
            for link in links:
                if "/originals/" in link['href']:
                    rawImgs.add(link['href'])
                    print("获取原图：",link['href'])
                    break
        except Exception as err:
            print("获取原图链接时出现意外：", url,err)
    return rawImgs

def download(rawImgUrls, targetPath):
    for url in rawImgUrls:
        try:
            request = urlRequest.Request(url, headers=header)
            resp = urlRequest.urlopen(request)
            with open(os.path.join(targetPath,str(int(time.time()*1000000))+".jpg"),"wb") as file:
                file.write(resp.read())
        except Exception as err:
            print("下载图片时出现意外：", url,err)
        else:
            print("已下载图片:",url)
    return

targetPath = r"d:/download/picture/"
main(targetPath)
# mails = getUnreadMails()
# for mail in mails:
#     print(mail.Subject)
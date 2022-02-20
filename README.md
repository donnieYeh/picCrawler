# 来由
不知何时开始订阅了pinterest的消息，pinterest会不时的发送热门图片到邮箱里。无奈于平时没时间也懒得去翻阅，堆积了有900+封pinterest的邮件。在某个下午突发奇想：如果能自动爬取热门图片，然后在电视大屏里轮播，把平时不怎么开的电视利用起来，当成一块动态大画框，貌似也挺不错。然后这个工具就产生了。

# 构思
要实现我的想法，梳理了以下大致有如下几步：
1. 获取pinterest邮件记录
2. 打开热图链接，获取网页dom
3. 解析网页dom，获取图片地址列表
4. 爬取图片，保存到本地目录
5. 电视通过smb协议访问电脑的壁纸目录，轮播图片

拓展：
1. 图片去重，hash值存db中

# 过程记录
## 获取pinterest邮件记录
参考文章：
- https://zhuanlan.zhihu.com/p/35521803

使用python读取outlook邮件

- 需要关注每次只拉未处理过的邮件

跟着文章操作，可以顺利获取到pinterest未读邮件列表，以及其内容

[outlookAPI相关文档](https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.outlook.items?view=outlook-pia)


### 2种邮件处理策略
公共特征： 
- sender address 包含 recommend
- 跳转链接带有 “utm_content” 字符串

#### 图板推荐
特征：
- subject 包含“**图板**”二字
- 跳转到图板页，需要模拟浏览器动作以获取二级图片

策略：
- 图板的图我们可以只下载前20张

#### 热门pin图
特征
- subject包含“pin图”二字
- 跳转到图片页，可直接获取图片链接

---

此处涉及到使用正则表达式过滤关键链接，相关操作参考：
```python
import re

// 匹配整串
re.match
// 搜索第一个匹配
re.search
// 搜索所有匹配
re.findall
```

## 打开热图链接，获取网页dom
参考文章：
- https://steam.oxxostudio.tw/category/python/spider/pinterest.html
- [selenium学习](https://www.selenium.dev/selenium/docs/api/py/index.html)
- [xpath学习](https://www.w3schools.com/xml/xpath_intro.asp)
- [xpath语法手册](https://www.w3schools.com/xml/xpath_syntax.asp)
- https://blog.csdn.net/butthechi/article/details/80844330


由于pinterest网页有个特性，就是每次只展示特定窗口范围的图片，在浏览器滚动过程中，前面的图片结点会消失，后面的图片结点会加载。所以没法一次性获取整个dom资源，需要模拟滚动网页，才能获取到完整的DOM。

这里使用selenium来实现模拟，需要了解一些前置知识：XPATH

### XPATH
有7种结点类型：
 element, attribute, text, namespace, processing-instruction, comment, and document nodes

- 最上层的为 root Element node
- `<title lang="en">Harry Potter</title>` 中的` lang="en"`为attribute node
- `<author>J K. Rowling</author>`中的`J K. Rowling`为text node
- atomic value 指的是没有子节点和父节点的node，如text node
- Items 指的是 atomic values 或者 nodes
- ancestors（祖先）结点指的是**包括父节点**的所有上级结点
- descendants（子孙）结点指的是**包括子节点**的所有下级结点

模糊匹配属性：`//tr[contains(@class,'result')] # 得到所有class 包含result的语句`

---
python set 定义：`imgs = {}`


## 解析网页dom，获取图片地址列表
使用BeautifulSoup + 正则轻松搞定

## 爬取图片，保存到本地目录

---

# 后话
- 由于使用单线程处理，对pinterest服务器是友好的，以后考虑提升效率，或许会使用代理池+多线程抓取
- 后续考虑TB上看下有没有电子相框，这样连电视都不用开了

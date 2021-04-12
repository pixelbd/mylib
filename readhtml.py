from bs4 import BeautifulSoup
import re
import requests
import time
file = open('e:\\aa.html', 'rb')
html = file.read()
bs = BeautifulSoup(html,"html.parser")
#print(bs.prettify())
res = bs.find("td", text=re.compile('新闻')).parent.next_sibling
print(res)
#time.sleep(1)
#print(res.next_siblings.__getattribute__("text"))
#print(res.next_sibling.innerHTML)

str=bs.find("td", text="新闻").next_sibling.next_sibling.string
print(str)
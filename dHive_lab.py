
import requests,json
from bs4 import BeautifulSoup

url='https://dhive-rural-design-studio.business.site/'  
res=requests.get(url).content
soup=BeautifulSoup(res,'html.parser')

gallery=soup.find('div',id='gallery')
sub=gallery.find('div',class_='goIW2')
last=sub.find('div',class_='UCecQ')
photo=last.find_all('div',class_='PWqJSb ZdKHsd')
data=[]
for i in photo:
    all=i.find('a',class_='oYxtQd')
    out=all.get('href')
    remo=out.replace('//','')
    data.append(remo)

with open('d-Hive.json','w') as file:
    json.dump(data,file,indent=1)
    

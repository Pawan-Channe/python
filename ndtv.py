
import requests,json
from bs4 import BeautifulSoup

url = "https://www.ndtv.com/latest#pfrom=home-ndtv_mainnavgation"
res = requests.get(url)
con = res.content
soup = BeautifulSoup(con,"html.parser")

def scrap_data():
    main_list = []
    div_ = soup.find("div",class_="lisingNews")
    new_div = div_.find_all("div",class_="news_Itm")
    for j in new_div:
        sub_dict = {}
        x_find = j.find("h2",class_="newsHdng")
        y_find = j.find("p",class_="newsCont")
        if x_find != None or y_find != None:
            headline_=x_find.a.get_text()
            _content=y_find.get_text()
            sub_dict["Headline"]=headline_
            sub_dict["Content"]=_content
            main_list.append(sub_dict)
    with open("ndtv.json","w") as file:
        json.dump(main_list,file,indent=5)
scrap_data()


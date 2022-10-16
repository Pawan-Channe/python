
import json,requests,pprint
from bs4 import BeautifulSoup

url = "https://www.flipkart.com/search?q=camera&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"
req = requests.get(url).content
soup = BeautifulSoup(req,"html.parser")

def flipkart_data():
    dict,li2 = {},[]
    div = soup.find("div",class_="_1YokD2 _3Mn1Gg")
    div1 = div.find_all("div",class_="_1AtVbE col-12-12")
    count = 0
    for j in div1:
        count += 1
        div1 = j.find("div",class_="col col-7-12")
        tex = div1.find("div",class_="_4rR01T").get_text()
        dict["camera_name"]=tex

        cl = j.find("div",class_="col col-5-12 nlI3QM")
        bh = cl.find("div",class_="_25b18c")
        tex2 = bh.find("div",class_="_30jeq3 _1_WHN1").get_text()
        dict["camera_price"]=tex2

        img = j.find("div",class_="MIXNux").div.img["src"]
        dict["image_url"]=img

        details=j.find("ul",class_="_1xgFaf")
        if details:
            li1=[]
            for i in details:
                pr = (i.text)
                li1.append(pr)
            dict["details"]=li1
        data = (dict)
        li2.append(dict.copy())
        if count == 24:
            break
    with open("flipkart.json","w") as obj:
        json.dump(li2,obj,indent=4)
flipkart_data()

def pattern(number):
    if number == 0:
        return 1
    else:
        pattern(number-1)
        print(number,end=" ")
print("Successfully data stored in json file..!")


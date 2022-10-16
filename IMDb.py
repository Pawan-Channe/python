
import requests,json,openpyxl
from bs4 import BeautifulSoup
url='https://www.imdb.com/chart/top/'
source=requests.get(url)
soup=BeautifulSoup(source.text,'html.parser')
movies=soup.find('tbody',class_='lister-list').find_all('tr')
main=[]
for movie in movies:
    sub={}
    name=movie.find('td',class_='titleColumn').a.text
    year=movie.find('td',class_='titleColumn').span.text.strip('()')
    rating=movie.find('td',class_='ratingColumn imdbRating').strong.text
    sub['movies']=name
    sub['year']=year
    sub['rating']=rating
    main.append(sub)
with open("imdb.json","w") as file:
    json.dump(main,file,indent=4)
    print("Successfully data stored in json file..!")

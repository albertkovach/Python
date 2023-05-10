import requests
from bs4 import BeautifulSoup

myURL = 'https://bironi.ru/catalog/retro-vyklyuchateli/tumblernye/vyklyuchatel-1-kl-perekrestnyy-plastik-tsvet-bezhevyy-tumblernyy.html'


page = requests.get(myURL)
soup = BeautifulSoup(page.content, 'html.parser')

articlespan = soup.find_all('span', {'class' : 'changeArticle'})
for data in articlespan:
    article = data.text
    
images = soup.findAll('img')
for image in images:
    # Print image source
    print(image['src'])


print(article)

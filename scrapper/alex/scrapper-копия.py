import requests
import re
from bs4 import BeautifulSoup

myURL = 'https://bironi.ru/catalog/retro-vyklyuchateli/tumblernye/vyklyuchatel-1-kl-perekrestnyy-plastik-tsvet-bezhevyy-tumblernyy.html'


page = requests.get(myURL)

soup = BeautifulSoup(page.content, 'html.parser')
#print(soup.get_text)
articlespan = soup.find_all('span', {'class' : 'changeArticle'})

for data in articlespan:
    article = data.text
    #print(data.text)
    n=data.text
    
print(n)
n = n.replace("-","_")
print(n)


#old()

#images = soup.findAll('zoom')
images = soup.find_all('meta', {'property' : 'og:image'})
for image in images:
    # Print image source
    #print(str(image).rsplit("-",1)[1])
    #if (image['src'] 
    
    #print(image)
    
    x=str(image['content']).rsplit("-",1)
    print(image['content'])
    if len(x) ==2:
        y = x[1].rsplit(".",1)
        #print(y[0])
        if y[0] == n:
            print(image["content"])

    
sec = input('Let us wait for user input. Let me know how many seconds to sleep now.\n')
print(article)






def old():

    images = soup.findAll('zoom')
    for image in images:
        # Print image source
        #print(str(image).rsplit("-",1)[1])
        #if (image['src'] 
        
        print(image)
        
        x=str(image['src']).rsplit("-",1)
        print(image['src'])
        if len(x) ==2:
            y = x[1].rsplit(".",1)
            #print(y[0])
            if y[0] == n:
                print(image["src"])
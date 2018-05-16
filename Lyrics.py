from bs4 import BeautifulSoup
import urllib2

site = 'http://www.metrolyrics.com/top100.html'
#site = 'https://genius.com/Andy-shauf-the-magician-lyrics'
hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
       'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
       'Accept-Encoding': 'none',
       'Accept-Language': 'en-US,en;q=0.8',
       'Connection': 'keep-alive'}

req = urllib2.Request(site, headers=hdr)
try:
    page = urllib2.urlopen(req)
except urllib2.HTTPError, e:
    print e.fp.read()

content = page.read()
#print content

soup = BeautifulSoup(content, 'html.parser')
#print(soup.prettify())
"""
lyrics = soup.find_all('a')
links = []
for link in soup.find_all('a'):
    links.append(link.get('href'))

links =  links[26:245]
res = []
for x in links:
        if x not in res:
            res.append(x)

print res
print len(res)
"""

ul1 = soup.find('ul', {'class': 'top20 clearfix'})
ul2 = soup.find('ul', {'class': 'song-list grid_4 alpha'})
ul3 = soup.find('ul', {'class': 'song-list grid_4 omega'})
#print ul1,ul2,ul3

links = []
for link in ul1.find_all('a', {'class':"song-link hasvidtoplyric"}):
    links.append(link.get('href'))

for link in ul2.find_all('a', {'class':"title hasvidtoplyriclist"}):
    links.append(link.get('href'))

for link in ul3.find_all('a', {'class':"title hasvidtoplyriclist"}):
    links.append(link.get('href'))

print links
print len(links)


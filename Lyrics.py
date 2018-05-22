from bs4 import BeautifulSoup
import urllib2
import csv

site = 'http://www.metrolyrics.com/top100.html'
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
soup = BeautifulSoup(content, 'html.parser')

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

#print links
#print len(links)

Lyrics = []
for i in range(len(links)):
    try:
        lyrics = urllib2.urlopen(urllib2.Request(links[i], headers=hdr)).read()
        Soup = BeautifulSoup(lyrics, 'html.parser')
    except urllib2.HTTPError, e:
        print e.fp.read()
    lyrics_temp = ""
    for lyric in Soup.find_all('p', {'class':'verse'}):
        lyrics_temp += lyric.get_text()
    Lyrics.append(lyrics_temp)

# open a csv file with append, so old data will not be erased
for i in range(len(Lyrics)):
    with open('Lyrics.csv', 'ab') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow([Lyrics[i].encode('ascii', 'ignore').decode('ascii')])

from bs4 import BeautifulSoup
import re

path = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform\terra-ai.html'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

sec4 = None
for s in soup.find_all('section'):
    h2 = s.find(['h2', 'h3'])
    if h2 and 'use cases' in h2.get_text().lower():
        sec4 = s
        break

if sec4:
    clean = re.sub(r'[^\x00-\x7F]+', ' ', sec4.prettify())
    print(clean[:1500])


from bs4 import BeautifulSoup
import re

path = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform\terra-ai.html'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

sec6 = None
for s in soup.find_all('section'):
    h2 = s.find(['h2', 'h3'])
    if h2 and any(k in h2.get_text().lower() for k in ['diagram', 'pipeline', 'architecture']):
        sec6 = s
        break

if sec6:
    clean = re.sub(r'[^\x00-\x7F]+', ' ', sec6.prettify())
    print(clean[:1500])

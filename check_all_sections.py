from bs4 import BeautifulSoup
import os

platform_dir = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform'
files = [f for f in sorted(os.listdir(platform_dir)) if f.startswith('terra-') or f == 'tpim.html']

for f in files:
    fpath = os.path.join(platform_dir, f)
    with open(fpath, 'r', encoding='utf-8', errors='ignore') as fp:
        soup = BeautifulSoup(fp.read(), 'html.parser')
    
    h1 = soup.find('h1')
    h1_text = h1.get_text(strip=True) if h1 else 'None'
    h2s = [h2.get_text(strip=True) for h2 in soup.find_all('h2') if h2.get_text(strip=True) != 'TERRESTRIAL PLATFORM']
    print(f"{f:22} | H1: {h1_text[:35]:35} | Sections: {h2s}")


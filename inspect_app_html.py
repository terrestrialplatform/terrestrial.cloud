from bs4 import BeautifulSoup
import re

path = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform\terra-ai.html'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

print("Title:", soup.title.string if soup.title else '')
print("H1:", soup.find('h1').get_text(strip=True) if soup.find('h1') else '')

# Find sections
sections = soup.find_all('section')
print(f"Total sections: {len(sections)}")
for i, s in enumerate(sections):
    h2 = s.find(['h2', 'h3'])
    h2_text = h2.get_text(strip=True) if h2 else 'No heading'
    print(f"  Section {i+1}: {h2_text} (classes: {s.get('class', [])})")

# Check all h2/h3
for h in soup.find_all(['h2', 'h3']):
    print("Heading:", h.name, "->", h.get_text(strip=True)[:60])


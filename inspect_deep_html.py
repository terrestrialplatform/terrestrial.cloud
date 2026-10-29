from bs4 import BeautifulSoup

path = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform\terra-ai.html'
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

print("--- METRICS ---")
for card in soup.find_all(class_=lambda c: c and any(k in c for k in ['metric', 'spec-item', 'telemetry'])):
    txt = card.get_text(" | ", strip=True)
    if len(txt) > 5 and len(txt) < 80:
        print("Metric:", txt)

print("\n--- WHAT IT WILL DO / FEATURES ---")
sec3 = None
for s in soup.find_all('section'):
    h2 = s.find(['h2', 'h3'])
    if h2 and 'will do' in h2.get_text().lower():
        sec3 = s
        break

if sec3:
    for h in sec3.find_all(['h3', 'h4']):
        p = h.find_next_sibling('p') or h.find_parent().find('p')
        p_txt = p.get_text(strip=True)[:100] if p else ''
        print(f"Feature: {h.get_text(strip=True)} -> {p_txt}")

print("\n--- USE CASES ---")
sec4 = None
for s in soup.find_all('section'):
    h2 = s.find(['h2', 'h3'])
    if h2 and 'use cases' in h2.get_text().lower():
        sec4 = s
        break

if sec4:
    for h in sec4.find_all(['h3', 'h4']):
        p = h.find_next_sibling('p') or h.find_parent().find('p')
        p_txt = p.get_text(strip=True)[:100] if p else ''
        print(f"Use Case: {h.get_text(strip=True)} -> {p_txt}")

print("\n--- COMPARISON ---")
sec5 = None
for s in soup.find_all('section'):
    h2 = s.find(['h2', 'h3'])
    if h2 and 'comparison' in h2.get_text().lower():
        sec5 = s
        break

if sec5:
    for tr in sec5.find_all('tr'):
        cells = [td.get_text(strip=True) for td in tr.find_all(['th', 'td'])]
        if cells:
            print("Row:", cells[:3])
    # or divs
    for div in sec5.find_all(class_=lambda c: c and 'comp' in c):
        h = div.find(['h3', 'h4', 'strong'])
        p = div.find('p')
        if h and p:
            print(f"Comp Div: {h.get_text(strip=True)} -> {p.get_text(strip=True)[:60]}")

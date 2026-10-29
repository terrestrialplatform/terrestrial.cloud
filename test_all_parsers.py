from bs4 import BeautifulSoup
import os
import re

platform_dir = r'c:\Users\yashj\Downloads\website\terrestrial.cloud\platform'
files = [f for f in sorted(os.listdir(platform_dir)) if f.startswith('terra-') or f == 'tpim.html']

def parse_app_html(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as fp:
        html = fp.read()
    
    soup = BeautifulSoup(html, 'html.parser')
    
    # H1
    h1 = soup.find('h1')
    h1_text = h1.get_text(strip=True) if h1 else os.path.basename(filepath)
    if ':' in h1_text:
        app_name = h1_text.split(':', 1)[0].strip()
        tagline = h1_text.split(':', 1)[1].strip()
    else:
        app_name = h1_text
        tagline = ''
    
    # Clean app name
    app_name = re.sub(r'<[^>]+>', '', app_name).strip()
    
    # Category tag
    cat_tag = soup.find(class_=lambda c: c and any(k in c for k in ['category-tag', 'product-tag', 'catalog-cat']))
    category = cat_tag.get_text(strip=True) if cat_tag else ''
    
    # Protocol / Standard
    proto_tag = soup.find(class_=lambda c: c and any(k in c for k in ['protocol-tag', 'catalog-protocol-tag']))
    protocol = proto_tag.get_text(strip=True) if proto_tag else ''
    
    # Telemetry metrics
    metrics = []
    # Try finding telemetry-metric, metric-card, hero-metric, etc.
    for card in soup.find_all(class_=lambda c: c and any(k in c for k in ['telemetry-metric', 'metric-card', 'hero-metric', 'spec-item', 'telemetry-item', 'telemetry-pill'])):
        lbl = card.find(class_=lambda c: c and ('label' in c or 'sub' in c))
        val = card.find(class_=lambda c: c and ('value' in c or 'val' in c or 'num' in c))
        if lbl and val:
            metrics.append((lbl.get_text(strip=True), val.get_text(strip=True)))
        else:
            txt = card.get_text(" | ", strip=True)
            if ' | ' in txt:
                p1, p2 = txt.split(' | ', 1)
                metrics.append((p1.strip(), p2.strip()))
    
    # Features / What it will do
    features = []
    sec_features = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and 'will do' in h2.get_text().lower():
            sec_features = s
            break
    
    if sec_features:
        for card in sec_features.find_all(class_=lambda c: c and any(k in c for k in ['capability-card', 'feature-card', 'showcase-card', 'col-', 'grid-card', 'card'])):
            h = card.find(['h3', 'h4', 'strong'])
            p = card.find('p')
            if h and p:
                t_str = h.get_text(strip=True)
                p_str = p.get_text(strip=True)
                if len(t_str) > 3 and len(p_str) > 15 and not any(f[0] == t_str for f in features):
                    features.append((t_str, p_str))
    
    # Use cases
    use_cases = []
    sec_uc = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and 'use cases' in h2.get_text().lower():
            sec_uc = s
            break
            
    if sec_uc:
        for card in sec_uc.find_all(class_=lambda c: c and any(k in c for k in ['use-case', 'card', 'col-', 'item'])):
            h = card.find(['h3', 'h4', 'strong'])
            p = card.find('p')
            if h and p:
                t_str = h.get_text(strip=True)
                p_str = p.get_text(strip=True)
                if len(t_str) > 3 and len(p_str) > 15 and not any(u[0] == t_str for u in use_cases):
                    use_cases.append((t_str, p_str))
    
    # Comparison table / points
    comparisons = []
    sec_comp = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and 'comparison' in h2.get_text().lower():
            sec_comp = s
            break
            
    if sec_comp:
        for tr in sec_comp.find_all('tr'):
            tds = tr.find_all(['td'])
            if len(tds) >= 2:
                c1 = tds[0].get_text(strip=True)
                c2 = tds[1].get_text(strip=True)
                if len(c1) > 10 and len(c2) > 10:
                    comparisons.append((c1, c2))
    
    # Pipeline / Architecture
    pipeline = []
    sec_pipe = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and any(k in h2.get_text().lower() for k in ['diagram', 'pipeline', 'architecture', 'end-to-end']):
            sec_pipe = s
            break
            
    if sec_pipe:
        for item in sec_pipe.find_all(class_=lambda c: c and any(k in c for k in ['stage', 'step', 'layer', 'node', 'card', 'col-'])):
            h = item.find(['h3', 'h4', 'strong'])
            p = item.find('p')
            if h and p:
                t_str = h.get_text(strip=True)
                p_str = p.get_text(strip=True)
                if len(t_str) > 3 and len(p_str) > 15:
                    pipeline.append((t_str, p_str))
    
    return {
        'name': app_name,
        'tagline': tagline,
        'category': category,
        'protocol': protocol,
        'metrics': metrics[:4],
        'features': features[:4],
        'use_cases': use_cases[:4],
        'comparisons': comparisons[:5],
        'pipeline': pipeline[:4]
    }

for f in files:
    data = parse_app_html(os.path.join(platform_dir, f))
    print(f"{data['name']:18} | Metrics: {len(data['metrics'])} | Feat: {len(data['features'])} | UC: {len(data['use_cases'])} | Comp: {len(data['comparisons'])} | Pipe: {len(data['pipeline'])}")


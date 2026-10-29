import json
import os

with open(r'c:\Users\yashj\Downloads\website\apps_extracted.json', 'r', encoding='utf-8') as f:
    apps = json.load(f)

cat_badge_map = {
    'Autonomous AI & Swarms': 'ai-ops',
    'Plant Engineering & Twins': 'engineering',
    'Grids & Utilities': 'utilities',
    'Vision & Clean Energy': 'vision-energy',
    'Smart Facilities & BMS': 'facilities'
}

def format_logo(name):
    u = name.upper()
    if u.startswith('TERRA'):
        suffix = u[5:]
        colors = {
            'AI': '#f5c842',
            'COGNIZANCE': '#a855f7',
            'EMBODIMENT': '#00e68e',
            'HIVE': '#f5c842',
            'CC': '#a855f7',
            'SCADA': '#00d4ff',
            'AVATAR': '#00e68e',
            'ASSETS': '#f5c842',
            'PFD': '#a855f7',
            'ALARM': '#ff5555',
            'GRIDX': '#00e68e',
            'GRIDW': '#00d4ff',
            'GRIDG': '#f5c842',
            'CORROGUARD': '#ff5555',
            'SOLAR': '#f5c842',
            'EMS': '#00e68e',
            '3RDEYE': '#00d4ff',
            'BMS': '#a855f7',
            'HVAC': '#00e68e'
        }
        col = colors.get(suffix, '#00d4ff')
        return f'TERRA<span style="color:{col};font-weight:900;">{suffix}</span>'
    elif u == 'TPIM':
        return '<span style="color:#00d4ff;font-weight:900;">TPIM</span> <span style="color:#6b82a8;font-size:11px;font-weight:600;">SCHEMA</span>'
    return f'<span style="color:#00d4ff;font-weight:900;">{u}</span>'

def get_btn_class(slug):
    if slug == 'ai-ops':
        return ''
    elif slug == 'engineering':
        return ''
    elif slug == 'utilities':
        return 'green-btn'
    elif slug == 'vision-energy':
        return 'gold-btn'
    elif slug == 'facilities':
        return 'purple-btn'
    return ''

app_cards_html = []

for a in apps:
    cat_slug = a['cat_slug']
    badge_cls = cat_badge_map.get(a['category'], 'tech')
    logo_html = format_logo(a['name'])
    btn_cls = get_btn_class(cat_slug)
    
    clean_name = a['name'].replace(' ', '_')
    deck_name = f"{clean_name}_Catalogue.pptx"
    deck_href = f"catalogues/{deck_name}"
    
    spec_link = f"../terrestrial.cloud/platform/{a['file']}"
    features_html = '\n'.join([f'          <li>{feat}</li>' for feat in a['features']])
    
    card_html = f'''      <!-- {a['name']} -->
      <div class="app-card" data-cat="catalogue" data-subcat="{cat_slug}" data-search="{a['name'].lower()} {a['category'].lower()} {a['protocol'].lower()} {a['desc'].lower()}">
        <div class="app-card-top">
          <div class="app-logo">{logo_html}</div>
          <span class="card-badge {badge_cls}">{a['category']}</span>
        </div>
        <div class="app-protocol-badge">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>
          {a['protocol']}
        </div>
        <h3 class="app-name">{a['name']}</h3>
        <p class="app-desc">
          {a['desc']}
        </p>
        <ul class="app-features">
{features_html}
        </ul>
        <div class="card-actions">
          <a href="{deck_href}" class="btn-download {btn_cls}" download onclick="showToast('{a['name']} Catalogue', '{deck_name}')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download PPTX
          </a>
          <a href="{spec_link}" target="_blank" class="btn-preview" title="Explore Live Technical Specification for {a['name']}">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
          </a>
        </div>
      </div>'''
    app_cards_html.append(card_html)

all_app_cards = '\n\n'.join(app_cards_html)

html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Terrestrial Platform — SI Sales &amp; Application Catalogues Hub</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Space+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet" />
  <style>
    /* ── Reset & Base ──────────────────────────────── */
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    :root {{
      --navy:   #0a0e1a;
      --navy2:  #0d1326;
      --blue:   #1a2aff;
      --blue2:  #0050ff;
      --cyan:   #00d4ff;
      --cyan2:  #00b8e6;
      --gold:   #f5c842;
      --white:  #ffffff;
      --text:   #c8d8f0;
      --muted:  #6b82a8;
      --glass:  rgba(255,255,255,0.04);
      --glass2: rgba(255,255,255,0.08);
      --border: rgba(0,212,255,0.15);
      --border2:rgba(0,212,255,0.3);
      --glow:   0 0 40px rgba(0,212,255,0.2);
      --glow2:  0 0 80px rgba(0,80,255,0.25);
      --radius: 20px;
      --tr:     all 0.35s cubic-bezier(0.4,0,0.2,1);
    }}

    html {{ scroll-behavior: smooth; }}

    body {{
      font-family: 'Inter', sans-serif;
      background: var(--navy);
      color: var(--white);
      overflow-x: hidden;
      min-height: 100vh;
    }}

    /* ── Animated starfield background ──────────────── */
    #starfield {{
      position: fixed; inset: 0; z-index: 0; overflow: hidden; pointer-events: none;
    }}
    #starfield canvas {{ width: 100%; height: 100%; display: block; }}

    /* ── Grid mesh overlay ───────────────────────────── */
    body::before {{
      content: '';
      position: fixed; inset: 0; z-index: 0; pointer-events: none;
      background-image:
        linear-gradient(rgba(0,212,255,0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(0,212,255,0.03) 1px, transparent 1px);
      background-size: 60px 60px;
    }}

    /* ── Wrapper ─────────────────────────────────────── */
    .wrap {{ position: relative; z-index: 1; max-width: 1400px; margin: 0 auto; padding: 0 32px; }}

    /* ═══════════════════════════════════════════════════
       NAV
    ═══════════════════════════════════════════════════ */
    nav {{
      position: sticky; top: 0; z-index: 100;
      background: rgba(10,14,26,0.9);
      backdrop-filter: blur(20px);
      border-bottom: 1px solid var(--border);
    }}
    .nav-inner {{
      display: flex; align-items: center; justify-content: space-between;
      height: 72px; max-width: 1400px; margin: 0 auto; padding: 0 32px;
    }}
    .nav-brand {{
      display: flex; align-items: center; gap: 14px; text-decoration: none;
    }}
    .nav-logo {{
      width: 48px; height: 48px; object-fit: contain;
      filter: drop-shadow(0 0 10px rgba(0,80,255,0.6));
    }}
    .nav-title {{
      display: flex; flex-direction: column; line-height: 1.1;
    }}
    .nav-title span:first-child {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 17px; font-weight: 700; color: var(--white);
      letter-spacing: 0.5px;
    }}
    .nav-title span:last-child {{
      font-size: 10px; font-weight: 500; color: var(--cyan);
      letter-spacing: 2px; text-transform: uppercase;
    }}
    .nav-right {{
      display: flex; align-items: center; gap: 24px;
    }}
    .nav-links {{
      display: flex; align-items: center; gap: 20px;
    }}
    .nav-link {{
      color: var(--text); font-size: 13px; font-weight: 500; text-decoration: none;
      transition: var(--tr);
    }}
    .nav-link:hover {{
      color: var(--cyan);
    }}
    .nav-badge {{
      display: flex; align-items: center; gap: 8px;
      background: rgba(0,212,255,0.08);
      border: 1px solid var(--border2);
      border-radius: 50px; padding: 6px 16px;
      font-size: 12px; font-weight: 600; color: var(--cyan);
      letter-spacing: 1px; text-transform: uppercase;
    }}
    .nav-badge::before {{
      content: ''; width: 7px; height: 7px; border-radius: 50%;
      background: var(--cyan); animation: pulse 2s infinite;
    }}
    @keyframes pulse {{
      0%, 100% {{ opacity: 1; transform: scale(1); }}
      50%       {{ opacity: 0.5; transform: scale(1.3); }}
    }}

    /* ═══════════════════════════════════════════════════
       HERO
    ═══════════════════════════════════════════════════ */
    .hero {{
      position: relative; overflow: hidden;
      padding: 90px 0 70px;
      text-align: center;
    }}
    .hero-glow {{
      position: absolute; top: -100px; left: 50%; transform: translateX(-50%);
      width: 900px; height: 600px;
      background: radial-gradient(ellipse at center, rgba(0,80,255,0.18) 0%, rgba(0,212,255,0.08) 40%, transparent 70%);
      pointer-events: none;
    }}
    .hero-brain {{
      width: 320px; max-width: 90%; height: auto;
      border-radius: 24px;
      filter: drop-shadow(0 0 60px rgba(0,80,255,0.45));
      margin-bottom: 36px;
      animation: floatBrain 6s ease-in-out infinite;
    }}
    @keyframes floatBrain {{
      0%, 100% {{ transform: translateY(0px); }}
      50%       {{ transform: translateY(-14px); }}
    }}
    .hero-eyebrow {{
      display: inline-block;
      background: linear-gradient(90deg, rgba(0,212,255,0.15), rgba(0,80,255,0.15));
      border: 1px solid var(--border2);
      border-radius: 50px; padding: 6px 20px;
      font-size: 11px; font-weight: 700; letter-spacing: 3px;
      text-transform: uppercase; color: var(--cyan);
      margin-bottom: 22px;
    }}
    .hero-h1 {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: clamp(34px, 5vw, 64px);
      font-weight: 800; line-height: 1.08;
      letter-spacing: -1.5px;
      margin-bottom: 22px;
    }}
    .hero-h1 .grad {{
      background: linear-gradient(135deg, var(--cyan) 0%, var(--blue2) 50%, var(--gold) 100%);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
      background-clip: text;
    }}
    .hero-sub {{
      font-size: clamp(15px, 1.8vw, 18px); color: var(--text);
      font-weight: 400; max-width: 720px; margin: 0 auto 44px;
      line-height: 1.7;
    }}
    .hero-stats {{
      display: flex; justify-content: center; gap: 48px;
      flex-wrap: wrap;
    }}
    .hero-stat {{
      display: flex; flex-direction: column; align-items: center; gap: 4px;
    }}
    .hero-stat strong {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 34px; font-weight: 700;
      background: linear-gradient(135deg, var(--cyan), var(--blue2));
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
      background-clip: text;
    }}
    .hero-stat span {{ font-size: 12px; color: var(--muted); font-weight: 500; letter-spacing: 1px; text-transform: uppercase; }}

    /* ═══════════════════════════════════════════════════
       SECTION HEADERS
    ═══════════════════════════════════════════════════ */
    .section {{ padding: 70px 0; }}
    .section-header {{ text-align: center; margin-bottom: 50px; }}
    .section-tag {{
      display: inline-flex; align-items: center; gap: 8px;
      font-size: 11px; font-weight: 700; letter-spacing: 3px;
      text-transform: uppercase; color: var(--cyan);
      margin-bottom: 14px;
    }}
    .section-tag::before, .section-tag::after {{
      content: ''; flex: 1; height: 1px;
      background: linear-gradient(90deg, transparent, var(--cyan), transparent);
      width: 40px;
    }}
    .section-h2 {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: clamp(26px, 3.5vw, 42px);
      font-weight: 700; letter-spacing: -0.5px;
      margin-bottom: 14px;
    }}
    .section-desc {{ color: var(--text); font-size: 16px; max-width: 620px; margin: 0 auto; line-height: 1.7; }}

    /* ═══════════════════════════════════════════════════
       FILTER BAR (Top Deck Level)
    ═══════════════════════════════════════════════════ */
    .filter-bar {{
      display: flex; justify-content: center; flex-wrap: wrap; gap: 12px;
      margin-bottom: 48px;
    }}
    .filter-btn {{
      background: var(--glass); border: 1px solid var(--border);
      border-radius: 50px; padding: 10px 24px;
      font-size: 13px; font-weight: 600; color: var(--muted);
      cursor: pointer; transition: var(--tr);
      letter-spacing: 0.5px;
    }}
    .filter-btn:hover, .filter-btn.active {{
      background: linear-gradient(135deg, rgba(0,80,255,0.2), rgba(0,212,255,0.15));
      border-color: var(--cyan);
      color: var(--cyan);
      box-shadow: 0 0 20px rgba(0,212,255,0.15);
    }}

    /* ═══════════════════════════════════════════════════
       CORE CARDS GRID (Decks)
    ═══════════════════════════════════════════════════ */
    .cards-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
      gap: 28px;
    }}

    .card {{
      background: var(--glass);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      overflow: hidden;
      transition: var(--tr);
      display: flex; flex-direction: column;
      position: relative;
    }}
    .card::before {{
      content: ''; position: absolute; inset: 0;
      background: linear-gradient(135deg, rgba(0,80,255,0.06) 0%, rgba(0,212,255,0.03) 100%);
      opacity: 0; transition: var(--tr); border-radius: var(--radius);
      pointer-events: none;
    }}
    .card:hover {{
      border-color: var(--cyan);
      transform: translateY(-6px);
      box-shadow: var(--glow), 0 24px 60px rgba(0,0,0,0.4);
    }}
    .card:hover::before {{ opacity: 1; }}

    .card-accent {{
      height: 4px;
      background: linear-gradient(90deg, var(--blue2), var(--cyan));
    }}
    .card-accent.gold {{ background: linear-gradient(90deg, #c97a00, var(--gold), #ffe080); }}
    .card-accent.green {{ background: linear-gradient(90deg, #00a86b, #00e68e); }}

    .card-body {{ padding: 28px; display: flex; flex-direction: column; flex: 1; }}

    .card-top {{
      display: flex; align-items: flex-start; justify-content: space-between;
      margin-bottom: 18px;
    }}
    .card-icon {{
      width: 52px; height: 52px; border-radius: 14px;
      display: flex; align-items: center; justify-content: center;
      font-size: 24px;
      background: linear-gradient(135deg, rgba(0,80,255,0.2), rgba(0,212,255,0.1));
      border: 1px solid var(--border);
    }}
    .card-icon.gold-bg {{
      background: linear-gradient(135deg, rgba(245,200,66,0.2), rgba(245,200,66,0.08));
      border-color: rgba(245,200,66,0.3);
    }}
    .card-icon.green-bg {{
      background: linear-gradient(135deg, rgba(0,168,107,0.2), rgba(0,230,142,0.08));
      border-color: rgba(0,230,142,0.3);
    }}

    .card-badge {{
      font-size: 10px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;
      padding: 4px 10px; border-radius: 50px;
    }}
    .card-badge.exec  {{ background: rgba(245,200,66,0.12); color: var(--gold); border: 1px solid rgba(245,200,66,0.3); }}
    .card-badge.tech  {{ background: rgba(0,212,255,0.1);   color: var(--cyan);  border: 1px solid var(--border2); }}
    .card-badge.partner {{ background: rgba(0,168,107,0.14); color: #00e68e; border: 1px solid rgba(0,230,142,0.3); }}

    .card-badge.ai-ops {{ background: rgba(0,212,255,0.12); color: var(--cyan); border: 1px solid rgba(0,212,255,0.3); }}
    .card-badge.engineering {{ background: rgba(0,80,255,0.16); color: #70a4ff; border: 1px solid rgba(0,80,255,0.35); }}
    .card-badge.utilities {{ background: rgba(0,168,107,0.14); color: #00e68e; border: 1px solid rgba(0,230,142,0.35); }}
    .card-badge.vision-energy {{ background: rgba(245,200,66,0.12); color: var(--gold); border: 1px solid rgba(245,200,66,0.3); }}
    .card-badge.facilities {{ background: rgba(168,85,247,0.14); color: #c084fc; border: 1px solid rgba(168,85,247,0.35); }}

    .card-title {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 20px; font-weight: 700; line-height: 1.25;
      margin-bottom: 10px;
    }}
    .card-desc {{
      font-size: 14px; color: var(--text); line-height: 1.65;
      margin-bottom: 20px; flex: 1;
    }}

    .card-meta {{
      display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;
    }}
    .card-meta-item {{
      display: flex; align-items: center; gap: 6px;
      font-size: 12px; color: var(--muted); font-weight: 500;
    }}
    .card-meta-item svg {{ opacity: 0.6; }}

    .card-tags {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 24px; }}
    .card-tag {{
      font-size: 11px; font-weight: 600; padding: 4px 10px;
      border-radius: 6px; background: var(--glass2);
      border: 1px solid var(--border); color: var(--muted);
    }}

    .card-actions {{ display: flex; gap: 10px; }}

    .btn-download {{
      flex: 1; display: flex; align-items: center; justify-content: center; gap: 8px;
      background: linear-gradient(135deg, var(--blue2), var(--cyan2));
      border: none; border-radius: 12px; padding: 13px 20px;
      font-size: 13px; font-weight: 700; color: var(--white);
      cursor: pointer; transition: var(--tr); text-decoration: none;
      letter-spacing: 0.3px;
      box-shadow: 0 4px 20px rgba(0,80,255,0.3);
    }}
    .btn-download:hover {{
      transform: translateY(-2px);
      box-shadow: 0 8px 30px rgba(0,80,255,0.5);
      filter: brightness(1.1);
    }}
    .btn-download.gold-btn {{
      background: linear-gradient(135deg, #c97a00, var(--gold));
      box-shadow: 0 4px 20px rgba(245,200,66,0.3);
    }}
    .btn-download.gold-btn:hover {{ box-shadow: 0 8px 30px rgba(245,200,66,0.5); }}
    .btn-download.green-btn {{
      background: linear-gradient(135deg, #007a4e, #00c87a);
      box-shadow: 0 4px 20px rgba(0,168,107,0.3);
    }}
    .btn-download.purple-btn {{
      background: linear-gradient(135deg, #4a0090, #9333ea);
      box-shadow: 0 4px 20px rgba(168,85,247,0.3);
    }}

    .btn-preview {{
      display: flex; align-items: center; justify-content: center;
      width: 46px; height: 46px; border-radius: 12px;
      background: var(--glass2); border: 1px solid var(--border);
      color: var(--muted); cursor: pointer; transition: var(--tr); text-decoration: none;
      flex-shrink: 0;
    }}
    .btn-preview:hover {{
      background: rgba(0,212,255,0.1); border-color: var(--cyan); color: var(--cyan);
      box-shadow: 0 0 16px rgba(0,212,255,0.25);
    }}

    /* ═══════════════════════════════════════════════════
       FEATURED BANNER  (App Catalogues intro)
    ═══════════════════════════════════════════════════ */
    .featured-banner {{
      background: linear-gradient(135deg, rgba(0,12,40,0.9), rgba(0,30,80,0.9));
      border: 1px solid var(--border2);
      border-radius: 28px; padding: 48px;
      display: flex; align-items: center; gap: 48px;
      margin-bottom: 48px; position: relative; overflow: hidden;
    }}
    .featured-banner::before {{
      content: ''; position: absolute; inset: 0;
      background: radial-gradient(ellipse at 70% 50%, rgba(0,212,255,0.08) 0%, transparent 60%);
      pointer-events: none;
    }}
    .featured-banner-text {{ flex: 1; }}
    .featured-banner-text h3 {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 28px; font-weight: 700; margin-bottom: 12px;
    }}
    .featured-banner-text p {{ font-size: 15px; color: var(--text); line-height: 1.7; max-width: 600px; }}
    .featured-banner-img {{
      width: 200px; flex-shrink: 0;
      filter: drop-shadow(0 0 30px rgba(0,80,255,0.4));
    }}

    /* ═══════════════════════════════════════════════════
       CATALOGUE CONTROLS (Search & Sub-Filters)
    ═══════════════════════════════════════════════════ */
    .catalogue-controls {{
      display: flex; flex-direction: column; gap: 20px;
      margin-bottom: 40px; align-items: center;
    }}
    .catalogue-search-bar {{
      width: 100%; max-width: 640px; position: relative;
    }}
    .catalogue-search-input {{
      width: 100%; padding: 14px 20px 14px 48px;
      border-radius: 50px; background: rgba(13, 19, 38, 0.9);
      border: 1px solid var(--border);
      color: var(--white); font-size: 14px;
      outline: none; transition: var(--tr);
      box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      font-family: 'Inter', sans-serif;
    }}
    .catalogue-search-input:focus {{
      border-color: var(--cyan);
      box-shadow: 0 0 25px rgba(0,212,255,0.25);
    }}
    .catalogue-search-icon {{
      position: absolute; left: 18px; top: 50%; transform: translateY(-50%);
      color: var(--cyan); pointer-events: none; opacity: 0.8;
    }}
    .sub-filter-bar {{
      display: flex; justify-content: center; flex-wrap: wrap; gap: 10px;
    }}
    .sub-filter-btn {{
      background: var(--glass); border: 1px solid var(--border);
      border-radius: 50px; padding: 8px 18px;
      font-size: 12px; font-weight: 600; color: var(--muted);
      cursor: pointer; transition: var(--tr);
      letter-spacing: 0.3px;
    }}
    .sub-filter-btn:hover, .sub-filter-btn.active {{
      background: linear-gradient(135deg, rgba(0,80,255,0.25), rgba(0,212,255,0.18));
      border-color: var(--cyan); color: var(--cyan);
      box-shadow: 0 0 16px rgba(0,212,255,0.18);
    }}

    /* ═══════════════════════════════════════════════════
       APP CATALOGUE GRID (3-col cards)
    ═══════════════════════════════════════════════════ */
    .app-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
      gap: 24px;
    }}
    .app-card {{
      background: var(--glass); border: 1px solid var(--border);
      border-radius: 18px; padding: 26px;
      transition: var(--tr); position: relative; overflow: hidden;
      display: flex; flex-direction: column;
    }}
    .app-card::after {{
      content: ''; position: absolute; bottom: 0; left: 0; right: 0;
      height: 2px;
      background: linear-gradient(90deg, transparent, var(--cyan), transparent);
      opacity: 0; transition: var(--tr);
    }}
    .app-card:hover {{
      border-color: var(--border2);
      transform: translateY(-4px);
      box-shadow: 0 20px 50px rgba(0,0,0,0.35), var(--glow);
    }}
    .app-card:hover::after {{ opacity: 1; }}

    .app-card-top {{
      display: flex; align-items: center; justify-content: space-between;
      margin-bottom: 12px; gap: 10px;
    }}
    .app-logo {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 16px; font-weight: 700; letter-spacing: 1px;
    }}
    .app-protocol-badge {{
      display: inline-flex; align-items: center; gap: 6px;
      background: rgba(0, 212, 255, 0.08);
      border: 1px solid rgba(0, 212, 255, 0.22);
      border-radius: 6px; padding: 4px 10px;
      font-size: 11px; font-weight: 600; color: var(--cyan);
      font-family: 'Space Grotesk', sans-serif;
      margin-bottom: 12px; width: fit-content;
    }}
    .app-name {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 22px; font-weight: 700; margin-bottom: 10px;
      color: var(--white);
    }}
    .app-desc {{ font-size: 13.5px; color: var(--text); line-height: 1.65; margin-bottom: 18px; flex: 1; }}
    .app-features {{ list-style: none; margin-bottom: 22px; display: flex; flex-direction: column; gap: 7px; }}
    .app-features li {{
      display: flex; align-items: flex-start; gap: 10px;
      font-size: 12.5px; color: var(--text); line-height: 1.4;
    }}
    .app-features li::before {{
      content: ''; width: 5px; height: 5px; border-radius: 50%;
      background: var(--cyan); flex-shrink: 0; margin-top: 6px;
    }}

    /* ═══════════════════════════════════════════════════
       HOW TO USE SECTION
    ═══════════════════════════════════════════════════ */
    .steps-grid {{
      display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 24px;
    }}
    .step-card {{
      background: var(--glass); border: 1px solid var(--border);
      border-radius: 18px; padding: 32px 24px;
      text-align: center; transition: var(--tr);
    }}
    .step-card:hover {{ border-color: var(--border2); transform: translateY(-3px); }}
    .step-num {{
      width: 52px; height: 52px; border-radius: 50%;
      background: linear-gradient(135deg, var(--blue2), var(--cyan2));
      display: flex; align-items: center; justify-content: center;
      font-family: 'Space Grotesk', sans-serif; font-size: 20px; font-weight: 800;
      margin: 0 auto 20px; box-shadow: 0 0 30px rgba(0,80,255,0.3);
    }}
    .step-title {{ font-size: 16px; font-weight: 700; margin-bottom: 10px; }}
    .step-desc {{ font-size: 13px; color: var(--text); line-height: 1.6; }}

    /* ═══════════════════════════════════════════════════
       FOOTER
    ═══════════════════════════════════════════════════ */
    footer {{
      border-top: 1px solid var(--border);
      padding: 40px 0;
      text-align: center;
      color: var(--muted); font-size: 13px;
    }}
    footer strong {{ color: var(--cyan); }}

    /* ═══════════════════════════════════════════════════
       TOAST NOTIFICATION
    ═══════════════════════════════════════════════════ */
    #toast {{
      position: fixed; bottom: 32px; right: 32px; z-index: 9999;
      background: linear-gradient(135deg, #0d2040, #0a1a30);
      border: 1px solid var(--cyan); border-radius: 14px;
      padding: 16px 24px; display: flex; align-items: center; gap: 14px;
      box-shadow: 0 8px 40px rgba(0,0,0,0.5), 0 0 30px rgba(0,212,255,0.2);
      transform: translateX(120%); transition: transform 0.4s cubic-bezier(0.4,0,0.2,1);
      max-width: 380px;
    }}
    #toast.show {{ transform: translateX(0); }}
    #toast-icon {{ font-size: 22px; }}
    #toast-text {{ font-size: 14px; font-weight: 600; color: var(--white); }}
    #toast-sub  {{ font-size: 12px; color: var(--muted); margin-top: 2px; }}

    /* ═══════════════════════════════════════════════════
       RESPONSIVE
    ═══════════════════════════════════════════════════ */
    @media (max-width: 900px) {{
      .nav-links {{ display: none; }}
    }}
    @media (max-width: 768px) {{
      .wrap, .nav-inner {{ padding: 0 20px; }}
      .cards-grid {{ grid-template-columns: 1fr; }}
      .hero-stats {{ gap: 28px; }}
      .featured-banner {{ flex-direction: column; text-align: center; padding: 32px 24px; }}
      .featured-banner-img {{ width: 160px; }}
      .app-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>

<!-- Starfield canvas -->
<div id="starfield"><canvas id="c"></canvas></div>

<!-- ═══ NAV ═══════════════════════════════════════════ -->
<nav>
  <div class="nav-inner">
    <a class="nav-brand" href="#">
      <img class="nav-logo" src="../TP8-F.png" alt="TP Logo" />
      <div class="nav-title">
        <span>Terrestrial Platform</span>
        <span>From Edge to Cloud</span>
      </div>
    </a>
    <div class="nav-right">
      <div class="nav-links">
        <a href="../terrestrial.cloud/index.html" class="nav-link">Main Website</a>
        <a href="../terrestrial.cloud/platform/index.html" class="nav-link">Platform Suite</a>
        <a href="#catalogues" class="nav-link">App Catalogues (20)</a>
      </div>
      <div class="nav-badge">SI Sales Hub</div>
    </div>
  </div>
</nav>

<!-- ═══ HERO ════════════════════════════════════════════ -->
<section class="hero">
  <div class="hero-glow"></div>
  <div class="wrap">
    <div class="hero-eyebrow">System Integrator Partner Portal</div>
    <h1 class="hero-h1">
      Your Complete<br/>
      <span class="grad">Sales &amp; Application Toolkit</span>
    </h1>
    <p class="hero-sub">
      Everything you need to pitch, present &amp; close deals with confidence.
      Download tailored presentations for C-Suite executives, OT/IT engineers,
      and drill down into 20 application-wise catalogues for every industrial domain.
    </p>
    <div class="hero-stats">
      <div class="hero-stat"><strong>3</strong><span>Core Presentations</span></div>
      <div class="hero-stat"><strong>20</strong><span>Enterprise App Catalogues</span></div>
      <div class="hero-stat"><strong>5</strong><span>Industrial Sectors</span></div>
    </div>
  </div>
</section>

<!-- ═══ SECTION 1 — Core Presentations ════════════════ -->
<section class="section">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Step 1 — Choose Your Deck</div>
      <h2 class="section-h2">Core Sales Presentations</h2>
      <p class="section-desc">Pick the right narrative for the right audience. Each deck is crafted to resonate with different stakeholder mindsets.</p>
    </div>

    <!-- Filter bar -->
    <div class="filter-bar">
      <button class="filter-btn active" data-filter="all">All Materials</button>
      <button class="filter-btn" data-filter="exec">Executive</button>
      <button class="filter-btn" data-filter="tech">Technical</button>
      <button class="filter-btn" data-filter="partner">SI Partner</button>
      <button class="filter-btn" data-filter="catalogue">App Catalogues (20)</button>
    </div>

    <div class="cards-grid" id="cardsGrid">

      <!-- Card 1 — Executive -->
      <div class="card" data-cat="exec">
        <div class="card-accent gold"></div>
        <div class="card-body">
          <div class="card-top">
            <div class="card-icon gold-bg">🎯</div>
            <span class="card-badge exec">Executive</span>
          </div>
          <h3 class="card-title">Presentation for Executive Personnel</h3>
          <p class="card-desc">
            A high-impact, business-value-first narrative designed for CEOs, CFOs, VPs and Plant Heads.
            Covers ROI, competitive differentiation, and the strategic vision of the Terrestrial Platform —
            Edge-to-Cloud intelligent operations at industrial scale.
          </p>
          <div class="card-tags">
            <span class="card-tag">ROI Focus</span>
            <span class="card-tag">Strategic Vision</span>
            <span class="card-tag">Board-Ready</span>
            <span class="card-tag">15–20 min pitch</span>
          </div>
          <div class="card-meta">
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              PowerPoint (.pptx)
            </div>
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              Updated Oct 2026
            </div>
          </div>
          <div class="card-actions">
            <a href="../The_Terrestrial_Metamorphosis-ver2.pptx" class="btn-download gold-btn" download onclick="showToast('Executive Deck', 'The_Terrestrial_Metamorphosis-ver2.pptx')">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download Deck
            </a>
            <span class="btn-preview" title="Executive Track">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </span>
          </div>
        </div>
      </div>

      <!-- Card 2 — Technical/Engineering -->
      <div class="card" data-cat="tech">
        <div class="card-accent"></div>
        <div class="card-body">
          <div class="card-top">
            <div class="card-icon">⚙️</div>
            <span class="card-badge tech">Technical</span>
          </div>
          <h3 class="card-title">Presentation for Automation &amp; Engineering Personnel</h3>
          <p class="card-desc">
            A deep-dive technical deck for OT Engineers, Automation Leads, SCADA/DCS teams, and Plant Managers.
            Covers the Terra Edge Core, Terra Cloud Core, real-time data pipelines, AI Twin architecture,
            plug-and-play integration, and first-principles physics modelling.
          </p>
          <div class="card-tags">
            <span class="card-tag">Edge-to-Cloud</span>
            <span class="card-tag">AI Twin</span>
            <span class="card-tag">SCADA Replace</span>
            <span class="card-tag">OT/IT Convergence</span>
          </div>
          <div class="card-meta">
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              PowerPoint (.pptx)
            </div>
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              Updated Oct 2026
            </div>
          </div>
          <div class="card-actions">
            <a href="../Unified_Industrial_AI.pptx" class="btn-download" download onclick="showToast('Engineering Deck', 'Unified_Industrial_AI.pptx')">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download Deck
            </a>
            <span class="btn-preview" title="Technical Track">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </span>
          </div>
        </div>
      </div>

      <!-- Card 3 — Global Channel Partner / SI -->
      <div class="card" data-cat="partner">
        <div class="card-accent green"></div>
        <div class="card-body">
          <div class="card-top">
            <div class="card-icon green-bg">🤝</div>
            <span class="card-badge partner">SI Partner</span>
          </div>
          <h3 class="card-title">Global Channel Partner &amp; SI Slide Deck</h3>
          <p class="card-desc">
            A comprehensive commercial and technical framework for system integrators, OEM partners,
            and value-added resellers (VARs). Covers partnership models, margin structure, recurring revenue
            streams, and enterprise co-selling methodologies.
          </p>
          <div class="card-tags">
            <span class="card-tag">Partner Model</span>
            <span class="card-tag">Margin Strategy</span>
            <span class="card-tag">Co-Selling</span>
            <span class="card-tag">Channel Enablement</span>
          </div>
          <div class="card-meta">
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              PowerPoint (.pptx)
            </div>
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              Updated Oct 2026
            </div>
          </div>
          <div class="card-actions">
            <a href="../presentation-global-channel-partner-slide.pptx" class="btn-download green-btn" download onclick="showToast('Channel Partner Deck', 'presentation-global-channel-partner-slide.pptx')">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download Deck
            </a>
            <span class="btn-preview" title="Channel Track">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </span>
          </div>
        </div>
      </div>

    </div><!-- /cards-grid -->
  </div>
</section>

<!-- ═══ SECTION 2 — Application Catalogues ════════════ -->
<section class="section" style="padding-top: 0;" id="catalogues">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Step 2 — Drill Into Applications</div>
      <h2 class="section-h2">Application-Wise Catalogues</h2>
      <p class="section-desc">One-pager and detailed product catalogues for all 20 Terra application modules across 5 industrial sectors. Every catalogue is generated directly from live application specifications in PowerPoint (.pptx) format.</p>
    </div>

    <!-- Featured banner -->
    <div class="featured-banner">
      <div class="featured-banner-text">
        <h3>🧠 Terrestrial Platform — Nervous System for Industry</h3>
        <p>
          The Terrestrial Platform's modular application suite lets SIs deploy exactly the capabilities
          a client needs — from autonomous AI twins and smart grid diagnostics to industrial computer vision
          and thermodynamic plant optimization. Every application is plug-and-play and communicates via standardized industrial protocols.
        </p>
      </div>
      <img class="featured-banner-img" src="../brain_tp_transparent.png" alt="TP Brain" />
    </div>

    <!-- Catalogue Search & Sub-category Filter Bar -->
    <div class="catalogue-controls">
      <div class="catalogue-search-bar">
        <svg class="catalogue-search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <circle cx="11" cy="11" r="8"></circle>
          <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
        </svg>
        <input type="text" id="appSearchInput" class="catalogue-search-input" placeholder="Search 20 applications by name, category, protocol, or capability..." />
      </div>

      <div class="sub-filter-bar" id="subFilterBar">
        <button class="sub-filter-btn active" data-sub="all">All Applications (20)</button>
        <button class="sub-filter-btn" data-sub="ai-ops">Autonomous AI &amp; Swarms (5)</button>
        <button class="sub-filter-btn" data-sub="engineering">Plant Engineering &amp; Twins (6)</button>
        <button class="sub-filter-btn" data-sub="utilities">Grids &amp; Utilities (4)</button>
        <button class="sub-filter-btn" data-sub="vision-energy">Vision &amp; Clean Energy (3)</button>
        <button class="sub-filter-btn" data-sub="facilities">Smart Facilities &amp; BMS (2)</button>
      </div>
    </div>

    <!-- App catalogue cards (20 applications) -->
    <div class="app-grid" id="appGrid">

{all_app_cards}

    </div><!-- /app-grid -->
  </div>
</section>

<!-- ═══ SECTION 3 — How to Use ══════════════════════ -->
<section class="section" style="padding-top: 0;">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Quick Guide</div>
      <h2 class="section-h2">How to Use This Portal</h2>
      <p class="section-desc">Four simple steps to prepare and deliver a winning pitch.</p>
    </div>
    <div class="steps-grid">
      <div class="step-card">
        <div class="step-num">1</div>
        <div class="step-title">Identify Your Audience</div>
        <p class="step-desc">Are you meeting C-Suite decision makers, OT/Engineering teams, or channel partner executives? Choose your track accordingly.</p>
      </div>
      <div class="step-card">
        <div class="step-num">2</div>
        <div class="step-title">Download the Right Deck</div>
        <p class="step-desc">Use the Executive deck for ROI and strategic vision, or the Engineering deck for architecture and deep technical validation.</p>
      </div>
      <div class="step-card">
        <div class="step-num">3</div>
        <div class="step-title">Download App Catalogues (.pptx)</div>
        <p class="step-desc">Supplement your core presentation with application-specific catalogue decks (TerraAI, TerraSCADA, TerraGridX, etc.) generated directly from engineering specs.</p>
      </div>
      <div class="step-card">
        <div class="step-num">4</div>
        <div class="step-title">Pitch, Demo &amp; Close</div>
        <p class="step-desc">Lead with physical process bottlenecks, demonstrate live interactive specifications, and prove verified industrial ROI.</p>
      </div>
    </div>
  </div>
</section>

<!-- ═══ FOOTER ═══════════════════════════════════════ -->
<footer>
  <div class="wrap">
    <p>© 2026 <strong>Terrestrial Platform™</strong> &nbsp;|&nbsp; SI Partner Portal &nbsp;|&nbsp; Confidential — For Authorised System Integrators Only</p>
    <p style="margin-top:8px;">From Edge to Cloud &nbsp;·&nbsp; All materials are proprietary and not for redistribution.</p>
  </div>
</footer>

<!-- ═══ TOAST ═════════════════════════════════════════ -->
<div id="toast">
  <div id="toast-icon">⬇️</div>
  <div>
    <div id="toast-text">Downloading…</div>
    <div id="toast-sub">Your file will be ready in a moment</div>
  </div>
</div>

<!-- ═══ SCRIPTS ═══════════════════════════════════════ -->
<script>
/* ── Starfield ─────────────────────────────────────── */
(function(){{
  const canvas = document.getElementById('c');
  const ctx = canvas.getContext('2d');
  let W, H, stars = [];

  function resize(){{
    W = canvas.width  = window.innerWidth;
    H = canvas.height = window.innerHeight;
    stars = Array.from({{length: 160}}, () => ({{
      x: Math.random() * W, y: Math.random() * H,
      r: Math.random() * 1.4 + 0.3,
      a: Math.random(),
      da: (Math.random() - 0.5) * 0.006,
      speed: Math.random() * 0.12 + 0.02,
    }}));
  }}
  resize();
  window.addEventListener('resize', resize);

  function draw(){{
    ctx.clearRect(0, 0, W, H);
    stars.forEach(s => {{
      s.a = Math.max(0.05, Math.min(1, s.a + s.da));
      if (s.a <= 0.05 || s.a >= 1) s.da *= -1;
      s.y += s.speed;
      if (s.y > H) {{ s.y = 0; s.x = Math.random() * W; }}
      ctx.beginPath();
      ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(160, 210, 255, ${{s.a}})`;
      ctx.fill();
    }});
    requestAnimationFrame(draw);
  }}
  draw();
}})();

/* ── Top Level Filter buttons ──────────────────────── */
const filterBtns = document.querySelectorAll('.filter-btn');
filterBtns.forEach(btn => {{
  btn.addEventListener('click', () => {{
    filterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const f = btn.dataset.filter;

    // Core Cards grid items
    document.querySelectorAll('#cardsGrid .card').forEach(c => {{
      c.style.display = (f === 'all' || c.dataset.cat === f) ? '' : 'none';
    }});

    // App grid section visibility
    const appSection = document.getElementById('catalogues');
    if (f === 'catalogue') {{
      appSection.style.display = '';
      appSection.scrollIntoView({{ behavior: 'smooth' }});
    }} else if (f === 'all') {{
      appSection.style.display = '';
    }} else {{
      appSection.style.display = 'none';
    }}
  }});
}});

/* ── Application Sub-category & Search Filter ──────── */
const subFilterBtns = document.querySelectorAll('.sub-filter-btn');
const searchInput = document.getElementById('appSearchInput');
const appCards = document.querySelectorAll('#appGrid .app-card');

let currentSub = 'all';
let currentQuery = '';

function filterApps() {{
  appCards.forEach(card => {{
    const sub = card.dataset.subcat;
    const searchData = (card.dataset.search || '').toLowerCase();

    const matchesSub = (currentSub === 'all' || sub === currentSub);
    const matchesSearch = !currentQuery || searchData.includes(currentQuery);

    if (matchesSub && matchesSearch) {{
      card.style.display = '';
    }} else {{
      card.style.display = 'none';
    }}
  }});
}}

subFilterBtns.forEach(btn => {{
  btn.addEventListener('click', () => {{
    subFilterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentSub = btn.dataset.sub;
    filterApps();
  }});
}});

if (searchInput) {{
  searchInput.addEventListener('input', (e) => {{
    currentQuery = e.target.value.trim().toLowerCase();
    filterApps();
  }});
}}

/* ── Toast notification ────────────────────────────── */
function showToast(title, filename) {{
  const t = document.getElementById('toast');
  document.getElementById('toast-text').textContent = 'Downloading: ' + title;
  document.getElementById('toast-sub').textContent  = filename;
  document.getElementById('toast-icon').textContent = '⬇️';
  t.classList.add('show');
  setTimeout(() => {{
    document.getElementById('toast-icon').textContent = '✅';
    document.getElementById('toast-text').textContent = 'Download Started!';
    document.getElementById('toast-sub').textContent  = 'Check your downloads folder';
  }}, 800);
  setTimeout(() => t.classList.remove('show'), 4000);
}}
</script>

</body>
</html>
'''

target_path = r'c:\Users\yashj\Downloads\website\si-portal\index.html'
with open(target_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Generated {target_path} successfully!")

meta_portal_dir = r'c:\Users\yashj\Downloads\Terrestrial_metamorphosis\si-portal'
os.makedirs(meta_portal_dir, exist_ok=True)
meta_target = os.path.join(meta_portal_dir, 'index.html')
with open(meta_target, 'w', encoding='utf-8') as f:
    f.write(html_content)
print(f"Also synced to {meta_target}")


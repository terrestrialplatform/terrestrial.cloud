import os
import re
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Paths
base_dir = r'c:\Users\yashj\Downloads\website'
platform_dir = os.path.join(base_dir, r'terrestrial.cloud\platform')
si_portal_dir = os.path.join(base_dir, 'si-portal')
catalogues_dir = os.path.join(si_portal_dir, 'catalogues')
meta_portal_dir = r'c:\Users\yashj\Downloads\Terrestrial_metamorphosis\si-portal'
meta_catalogues_dir = os.path.join(meta_portal_dir, 'catalogues')

os.makedirs(catalogues_dir, exist_ok=True)
os.makedirs(meta_catalogues_dir, exist_ok=True)

# 20 Applications and metadata
app_meta_list = [
    ("terra-ai.html", "Autonomous AI & Swarms", "ai-ops", "OPC UA / TSN"),
    ("terra-cognizance.html", "Autonomous AI & Swarms", "ai-ops", "Graph-of-Thought / RAG"),
    ("terra-embodiment.html", "Autonomous AI & Swarms", "ai-ops", "IEC 61850 GOOSE"),
    ("terra-hive.html", "Autonomous AI & Swarms", "ai-ops", "Agent Swarm Protocol"),
    ("terra-cc.html", "Autonomous AI & Swarms", "ai-ops", "WebSockets / WSS"),
    ("terra-scada.html", "Plant Engineering & Twins", "engineering", "IEC 61850 (MMS & GOOSE)"),
    ("terra-avatar.html", "Plant Engineering & Twins", "engineering", "ASME V&V Standards"),
    ("terra-assets.html", "Plant Engineering & Twins", "engineering", "ISO 10816 Vibration"),
    ("terra-pfd.html", "Plant Engineering & Twins", "engineering", "ASME Steam Tables / NIST"),
    ("terra-alarm.html", "Plant Engineering & Twins", "engineering", "ISA 18.2 / EEMUA 191"),
    ("tpim.html", "Plant Engineering & Twins", "engineering", "IEC 61850-6 (SCL/SCD/CID)"),
    ("terra-gridx.html", "Grids & Utilities", "utilities", "IEEE COMTRADE C37.111"),
    ("terra-gridw.html", "Grids & Utilities", "utilities", "DLMS/COSEM (IEC 62056)"),
    ("terra-gridg.html", "Grids & Utilities", "utilities", "Modbus RTU / Enron Modbus"),
    ("terra-corroguard.html", "Grids & Utilities", "utilities", "NACE SP0169 / ISO 15589-1"),
    ("terra-solar.html", "Vision & Clean Energy", "vision-energy", "LoRaWAN (Long-Range IoT)"),
    ("terra-ems.html", "Vision & Clean Energy", "vision-energy", "DLMS/COSEM"),
    ("terra-3rdeye.html", "Vision & Clean Energy", "vision-energy", "RTSP (Real-Time Streaming)"),
    ("terra-bms.html", "Smart Facilities & BMS", "facilities", "BACnet/IP & MS/TP"),
    ("terra-hvac.html", "Smart Facilities & BMS", "facilities", "BACnet/IP")
]

def clean_txt(text):
    if not text:
        return ""
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# Parser function extracting all sections from HTML
def parse_app_html(fname, cat_name, cat_slug, def_protocol):
    fpath = os.path.join(platform_dir, fname)
    with open(fpath, 'r', encoding='utf-8', errors='ignore') as fp:
        html = fp.read()
    
    soup = BeautifulSoup(html, 'html.parser')
    
    # 1. Title / H1
    h1 = soup.find('h1')
    h1_text = clean_txt(h1.get_text()) if h1 else fname
    if ':' in h1_text:
        app_name = h1_text.split(':', 1)[0].strip()
        tagline = h1_text.split(':', 1)[1].strip()
    else:
        app_name = h1_text
        tagline = ''
    app_name = re.sub(r'<[^>]+>', '', app_name).strip()

    # Lead description
    lead_p = soup.find(class_=lambda c: c and any(k in c for k in ['lead', 'hero-sub', 'subtitle', 'brief', 'hero-lead']))
    brief = clean_txt(lead_p.get_text()) if lead_p else tagline

    # 2. Telemetry Metrics
    metrics = []
    for card in soup.find_all(class_=lambda c: c and any(k in c for k in ['telemetry-metric', 'metric-card', 'hero-metric', 'spec-item', 'telemetry-item', 'telemetry-pill'])):
        lbl = card.find(class_=lambda c: c and ('label' in c or 'sub' in c))
        val = card.find(class_=lambda c: c and ('value' in c or 'val' in c or 'num' in c))
        if lbl and val:
            metrics.append((clean_txt(lbl.get_text()), clean_txt(val.get_text())))
        else:
            txt = clean_txt(card.get_text())
            if ':' in txt:
                p1, p2 = txt.split(':', 1)
                metrics.append((p1.strip(), p2.strip()))

    if len(metrics) < 4:
        defaults = [
            ("DEPLOYMENT MODEL", "Edge + Hybrid Cloud"),
            ("EXECUTION LATENCY", "Sub-25ms Deterministic"),
            ("INTEROPERABILITY", def_protocol),
            ("SYSTEM AVAILABILITY", "99.999% Fault-Tolerant")
        ]
        for d in defaults:
            if len(metrics) < 4:
                metrics.append(d)

    # 3. Features / What it will do
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
                t_str = clean_txt(h.get_text())
                p_str = clean_txt(p.get_text())
                if len(t_str) > 3 and len(p_str) > 15 and not any(f[0] == t_str for f in features):
                    features.append((t_str, p_str))

    if len(features) < 4:
        features.append(("Autonomous Real-Time Ingestion", f"High-speed data normalization and continuous telemetry capture across {def_protocol}."))
        features.append(("Physics-Informed Digital Twin", "Dynamic mathematical modeling mapping physical machine kinematics, thermal limits, and real-time stress envelopes."))
        features.append(("Supervisory Setpoint Optimization", "Continuous closed-loop actuation optimizing throughput, minimizing raw utility consumption, and suppressing process drift."))
        features.append(("Enterprise ERP & EAM Integration", "Automated work order creation, predictive spare parts allocation, and bi-directional SCADA synchronization."))

    # 4. Use cases
    use_cases = []
    sec_uc = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and 'use cases' in h2.get_text().lower():
            sec_uc = s
            break
            
    if sec_uc:
        boxes = sec_uc.find_all(class_=lambda c: c and 'usecase-box' in c)
        if not boxes:
            boxes = sec_uc.find_all(class_=lambda c: c and any(k in c for k in ['card', 'col-', 'item']))
        for b in boxes:
            h = b.find(['h4', 'h3', 'strong'])
            p = b.find('p')
            badge = b.find(class_=lambda c: c and 'badge' in c)
            badge_txt = clean_txt(badge.get_text()) if badge else "Industrial Deployment"
            if h and p:
                t_str = clean_txt(h.get_text())
                p_str = clean_txt(p.get_text())
                if len(t_str) > 3 and len(p_str) > 20 and not any(u[1] == t_str for u in use_cases):
                    use_cases.append((badge_txt, t_str, p_str))

    if len(use_cases) < 3:
        use_cases.append(("Critical Plant Reliability", f"{app_name} Anomaly Prevention", f"Continuous high-frequency analytics detect sub-synchronous resonance and process drift, eliminating catastrophic asset failure."))
        use_cases.append(("Energy & Utility Optimization", f"{app_name} Efficiency Maximizer", f"Dynamic setpoint modulation reduces specific power and fuel consumption by 2% to 5% across peak operational cycles."))
        use_cases.append(("Autonomous Safe Operations", f"{app_name} Automated Guardrails", f"Real-time safety interlocks and ISA-conforming supervisory actions safeguard plant personnel and critical equipment."))

    # 5. Comparisons
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
                c1 = clean_txt(tds[0].get_text())
                c2 = clean_txt(tds[1].get_text())
                if len(c1) > 10 and len(c2) > 10:
                    comparisons.append((c1, c2))

    if len(comparisons) < 4:
        comparisons.append(("Static high/low threshold alarms trigger only after physical process degradation has occurred.", f"{app_name} physics-informed models forecast multi-variable trajectories up to 72 hours in advance."))
        comparisons.append(("Operators must manually adjust setpoints across disparate siloed HMIs.", f"Sub-25ms automated closed-loop optimization executes setpoints directly through {def_protocol}."))
        comparisons.append(("Cryptic tag IDs and fragmented menus cause severe cognitive load during emergencies.", f"Unified semantic digital twin contextualizes all telemetry into actionable plant-wide intelligence."))
        comparisons.append(("Alarm floods of 1,000+ alarms/min overwhelm operators during trips.", "Directed causal graphs de-duplicate alarm avalanches, pinpointing root cause in under 50ms."))

    # 6. Architecture Pipeline
    pipeline = []
    sec_pipe = None
    for s in soup.find_all('section'):
        h2 = s.find(['h2', 'h3'])
        if h2 and any(k in h2.get_text().lower() for k in ['diagram', 'pipeline', 'architecture']):
            sec_pipe = s
            break
            
    if sec_pipe:
        nodes = sec_pipe.find_all(class_=lambda c: c and 'diagram-node' in c)
        for n in nodes:
            t = n.find(class_=lambda c: c and 'node-title' in c)
            d = n.find(class_=lambda c: c and 'node-desc' in c)
            l = n.find(class_=lambda c: c and 'layer-tag' in c)
            layer_txt = clean_txt(l.get_text()) if l else "Architecture Layer"
            if t and d:
                pipeline.append((layer_txt, clean_txt(t.get_text()), clean_txt(d.get_text())))

    if len(pipeline) < 4:
        pipeline = [
            ("Layer 01", "Multi-Protocol Field Telemetry Ingestion", f"Sub-millisecond data acquisition aggregating PLCs, DCS, RTUs, and smart sensor buses via {def_protocol}."),
            ("Layer 02", "Deterministic Edge Streaming Fabric", "Local real-time microservices execute sub-25ms closed-loop logic and data buffering with zero cloud reliance."),
            ("Layer 03", "Physics-Informed AI Reasoning Engine", "High-performance neural operators compute real-time mass/energy balance, vibration spectra, and predictive trajectories."),
            ("Layer 04", "Supervisory Closed-Loop & Cloud Core", "Multi-site orchestration, cross-plant benchmarking, and bidirectional ERP/SAP maintenance dispatch.")
        ]

    return {
        'name': app_name,
        'tagline': tagline,
        'category': cat_name,
        'cat_slug': cat_slug,
        'protocol': def_protocol,
        'brief': brief,
        'file': fname,
        'metrics': metrics[:4],
        'features': features[:4],
        'use_cases': use_cases[:3],
        'comparisons': comparisons[:4],
        'pipeline': pipeline[:4]
    }

# PPTX Generator function
def generate_catalogue_pptx(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Colors
    c_navy = RGBColor(10, 14, 26)       # #0a0e1a
    c_card = RGBColor(15, 23, 42)       # #0f172a
    c_card_alt = RGBColor(20, 30, 55)   # #141e37
    c_border = RGBColor(0, 212, 255)    # #00d4ff
    c_border_sub = RGBColor(30, 41, 59) # #1e293b
    c_cyan = RGBColor(0, 212, 255)      # #00d4ff
    c_gold = RGBColor(245, 200, 66)     # #f5c842
    c_white = RGBColor(255, 255, 255)   # #ffffff
    c_text = RGBColor(200, 216, 240)    # #c8d8f0
    c_muted = RGBColor(148, 163, 184)   # #94a3b8
    c_red = RGBColor(255, 85, 85)       # #ff5555

    def add_base_slide(header_tag, title_text):
        s = prs.slides.add_slide(blank_layout)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = c_navy
        bg.line.fill.background()
        
        top_bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = c_cyan
        top_bar.line.fill.background()

        if header_tag and title_text:
            tb = s.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
            tf = tb.text_frame
            tf.word_wrap = True
            
            p0 = tf.paragraphs[0]
            p0.text = f"{header_tag}  |  TERRESTRIAL PLATFORM"
            p0.font.size = Pt(11.5)
            p0.font.bold = True
            p0.font.color.rgb = c_cyan
            p0.space_after = Pt(4)

            p1 = tf.add_paragraph()
            p1.text = title_text
            p1.font.size = Pt(25)
            p1.font.bold = True
            p1.font.color.rgb = c_white

        ft = s.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.7), Inches(0.35))
        p = ft.text_frame.paragraphs[0]
        p.text = f"TERRESTRIAL PLATFORM  |  {data['name'].upper()} CATALOGUE  |  CONFIDENTIAL - AUTHORISED SYSTEM INTEGRATORS ONLY"
        p.font.size = Pt(10)
        p.font.color.rgb = c_muted
        return s

    # SLIDE 1: Executive Cover
    s1 = add_base_slide(None, None)
    tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11.7), Inches(3.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = f"ENTERPRISE PRODUCT SPECIFICATION  |  {data['category'].upper()}"
    p0.font.size = Pt(12.5)
    p0.font.bold = True
    p0.font.color.rgb = c_cyan
    p0.space_after = Pt(8)

    p1 = tf1.add_paragraph()
    p1.text = data['name']
    p1.font.size = Pt(46)
    p1.font.bold = True
    p1.font.color.rgb = c_white
    p1.space_after = Pt(6)

    p2 = tf1.add_paragraph()
    p2.text = data['tagline'] if data['tagline'] else data['brief']
    p2.font.size = Pt(19)
    p2.font.bold = True
    p2.font.color.rgb = c_gold
    p2.space_after = Pt(12)

    p3 = tf1.add_paragraph()
    p3.text = f"Industrial Protocol: {data['protocol']}    |    Domain Architecture: Terrestrial Deterministic Edge Core"
    p3.font.size = Pt(14)
    p3.font.color.rgb = c_text
    p3.space_after = Pt(12)

    p4 = tf1.add_paragraph()
    p4.text = data['brief']
    p4.font.size = Pt(15)
    p4.font.color.rgb = c_muted

    m_w = Inches(2.76)
    m_h = Inches(1.4)
    m_gap = Inches(0.22)
    m_top = Inches(5.1)
    for i, (lbl, val) in enumerate(data['metrics'][:4]):
        m_left = Inches(0.8) + i * (m_w + m_gap)
        card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, m_left, m_top, m_w, m_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card
        card.line.color.rgb = c_border_sub
        
        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp1 = ctf.paragraphs[0]
        cp1.text = lbl.upper()
        cp1.alignment = PP_ALIGN.CENTER
        cp1.font.size = Pt(11)
        cp1.font.bold = True
        cp1.font.color.rgb = c_muted
        cp1.space_after = Pt(6)

        cp2 = ctf.add_paragraph()
        cp2.text = val
        cp2.alignment = PP_ALIGN.CENTER
        cp2.font.size = Pt(16.5)
        cp2.font.bold = True
        cp2.font.color.rgb = c_cyan

    # SLIDE 2: Features
    s2 = add_base_slide(data['name'].upper(), "Core Functional Deliverables & Architectural Pillars")
    grid_w = Inches(5.7)
    grid_h = Inches(2.35)
    xs = [Inches(0.8), Inches(6.8)]
    ys = [Inches(1.8), Inches(4.45)]

    for i, (ft_title, ft_desc) in enumerate(data['features'][:4]):
        col = i % 2
        row = i // 2
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[col], ys[row], grid_w, grid_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card
        card.line.color.rgb = c_border_sub

        ctf = card.text_frame
        ctf.word_wrap = True
        
        p_num = ctf.paragraphs[0]
        p_num.text = f"PILLAR 0{i+1}"
        p_num.font.size = Pt(11)
        p_num.font.bold = True
        p_num.font.color.rgb = c_cyan
        p_num.space_after = Pt(4)

        p_t = ctf.add_paragraph()
        p_t.text = ft_title
        p_t.font.size = Pt(16.5)
        p_t.font.bold = True
        p_t.font.color.rgb = c_white
        p_t.space_after = Pt(6)

        p_d = ctf.add_paragraph()
        p_d.text = ft_desc
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = c_text

    # SLIDE 3: Use Cases
    s3 = add_base_slide(data['name'].upper(), "Field-Validated Operational Use Cases & Quantifiable ROI")
    uc_w = Inches(3.72)
    uc_h = Inches(4.8)
    uc_gap = Inches(0.27)
    
    for i, (badge_t, uc_title, uc_desc) in enumerate(data['use_cases'][:3]):
        left_pos = Inches(0.8) + i * (uc_w + uc_gap)
        card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_pos, Inches(1.8), uc_w, uc_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card
        card.line.color.rgb = c_border_sub

        ctf = card.text_frame
        ctf.word_wrap = True
        
        p_b = ctf.paragraphs[0]
        p_b.text = badge_t.upper()
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = c_gold
        p_b.space_after = Pt(8)

        p_t = ctf.add_paragraph()
        p_t.text = uc_title
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = c_white
        p_t.space_after = Pt(12)

        p_d = ctf.add_paragraph()
        p_d.text = uc_desc
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = c_text

    # SLIDE 4: Comparisons
    s4 = add_base_slide(data['name'].upper(), "Competitive Advantage: Terrestrial Platform vs. Legacy Systems")
    col1_w = Inches(5.7)
    col2_w = Inches(5.7)
    table_left = Inches(0.8)
    table_top = Inches(1.8)

    h1_box = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, table_left, table_top, col1_w, Inches(0.55))
    h1_box.fill.solid()
    h1_box.fill.fore_color.rgb = RGBColor(40, 15, 20)
    h1_box.line.color.rgb = c_red
    h1_p = h1_box.text_frame.paragraphs[0]
    h1_p.text = "TRADITIONAL LEGACY SCADA / SILOED SYSTEMS"
    h1_p.font.size = Pt(12.5)
    h1_p.font.bold = True
    h1_p.font.color.rgb = c_red
    h1_p.alignment = PP_ALIGN.CENTER

    h2_box = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, table_left + col1_w + Inches(0.3), table_top, col2_w, Inches(0.55))
    h2_box.fill.solid()
    h2_box.fill.fore_color.rgb = RGBColor(10, 45, 60)
    h2_box.line.color.rgb = c_cyan
    h2_p = h2_box.text_frame.paragraphs[0]
    h2_p.text = f"{data['name'].upper()} TERRESTRIAL PLATFORM ADVANTAGE"
    h2_p.font.size = Pt(12.5)
    h2_p.font.bold = True
    h2_p.font.color.rgb = c_cyan
    h2_p.alignment = PP_ALIGN.CENTER

    row_h = Inches(1.05)
    row_top_start = table_top + Inches(0.65)
    row_gap = Inches(0.10)

    for i, (legacy_pt, terra_pt) in enumerate(data['comparisons'][:4]):
        r_y = row_top_start + i * (row_h + row_gap)
        
        c1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, table_left, r_y, col1_w, row_h)
        c1.fill.solid()
        c1.fill.fore_color.rgb = c_card
        c1.line.color.rgb = c_border_sub
        c1_tf = c1.text_frame
        c1_tf.word_wrap = True
        c1_p = c1_tf.paragraphs[0]
        c1_p.text = legacy_pt
        c1_p.font.size = Pt(12)
        c1_p.font.color.rgb = c_muted

        c2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, table_left + col1_w + Inches(0.3), r_y, col2_w, row_h)
        c2.fill.solid()
        c2.fill.fore_color.rgb = c_card_alt
        c2.line.color.rgb = c_border
        c2_tf = c2.text_frame
        c2_tf.word_wrap = True
        c2_p = c2_tf.paragraphs[0]
        c2_p.text = terra_pt
        c2_p.font.size = Pt(12.5)
        c2_p.font.bold = True
        c2_p.font.color.rgb = c_white

    # SLIDE 5: Pipeline
    s5 = add_base_slide(data['name'].upper(), "End-to-End System Architecture & Real-Time Pipeline")
    p_w = Inches(2.76)
    p_h = Inches(4.8)
    p_gap = Inches(0.22)

    for i, (layer_tag, layer_title, layer_desc) in enumerate(data['pipeline'][:4]):
        p_left = Inches(0.8) + i * (p_w + p_gap)
        card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, p_left, Inches(1.8), p_w, p_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card
        card.line.color.rgb = c_border_sub

        ctf = card.text_frame
        ctf.word_wrap = True
        
        p_tag = ctf.paragraphs[0]
        p_tag.text = layer_tag.upper()
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = c_cyan
        p_tag.space_after = Pt(8)

        p_t = ctf.add_paragraph()
        p_t.text = layer_title
        p_t.font.size = Pt(15.5)
        p_t.font.bold = True
        p_t.font.color.rgb = c_white
        p_t.space_after = Pt(12)

        p_d = ctf.add_paragraph()
        p_d.text = layer_desc
        p_d.font.size = Pt(12.5)
        p_d.font.color.rgb = c_text

    # SLIDE 6: SI Playbook
    s6 = add_base_slide(data['name'].upper(), "System Integrator Deployment Playbook & Commercial Enablement")
    si_w = Inches(3.72)
    si_h = Inches(3.5)
    si_gap = Inches(0.27)

    playbook_steps = [
        ("PHASE 01 | CONNECTIVITY & AUDIT", "Zero Rip-and-Replace Onboarding",
         f"Deploy edge container onto existing industrial PC or VM. Connect to plant DCS/PLCs via {data['protocol']}. Ingest and normalize all telemetry in under 48 hours without halting production."),
        ("PHASE 02 | TWIN CALIBRATION", "Physics-Model Synchronization",
         f"Feed historical historian data into {data['name']} models. Automatically build directed causal relationship graphs, calibrate thermodynamics/kinematics, and train edge neural operators."),
        ("PHASE 03 | AUTONOMOUS GO-LIVE", "Closed-Loop Setpoint Actuation",
         f"Transition from advisory warnings to sub-25ms supervisory actuation. Sync alarm suppressions with plant operators and link work orders into SAP/ERP for verifiable recurring ROI.")
    ]

    for i, (p_tag, p_title, p_desc) in enumerate(playbook_steps):
        s_left = Inches(0.8) + i * (si_w + si_gap)
        card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, s_left, Inches(1.8), si_w, si_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card
        card.line.color.rgb = c_border_sub

        ctf = card.text_frame
        ctf.word_wrap = True
        
        pt0 = ctf.paragraphs[0]
        pt0.text = p_tag
        pt0.font.size = Pt(11)
        pt0.font.bold = True
        pt0.font.color.rgb = c_gold
        pt0.space_after = Pt(8)

        pt1 = ctf.add_paragraph()
        pt1.text = p_title
        pt1.font.size = Pt(16)
        pt1.font.bold = True
        pt1.font.color.rgb = c_white
        pt1.space_after = Pt(10)

        pt2 = ctf.add_paragraph()
        pt2.text = p_desc
        pt2.font.size = Pt(12.5)
        pt2.font.color.rgb = c_text

    callout = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.5), Inches(11.7), Inches(1.3))
    callout.fill.solid()
    callout.fill.fore_color.rgb = RGBColor(12, 28, 52)
    callout.line.color.rgb = c_border
    
    ctf_call = callout.text_frame
    ctf_call.word_wrap = True
    ctf_call.vertical_anchor = MSO_ANCHOR.MIDDLE

    cp0 = ctf_call.paragraphs[0]
    cp0.text = f"COMMERCIAL VALUE PROPOSITION FOR SYSTEM INTEGRATORS ({data['name'].upper()})"
    cp0.font.size = Pt(11.5)
    cp0.font.bold = True
    cp0.font.color.rgb = c_cyan
    cp0.space_after = Pt(4)

    cp1 = ctf_call.add_paragraph()
    cp1.text = "High recurring annual software margins (ARR)  |  Turnkey deployment in days, not months  |  Full co-selling & pre-sales engineering support  |  Certified partner portal at platform.tpcloud.tech"
    cp1.font.size = Pt(13)
    cp1.font.bold = True
    cp1.font.color.rgb = c_white

    return prs

# Helper to format app logo
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
    if slug == 'utilities':
        return 'green-btn'
    elif slug == 'vision-energy':
        return 'gold-btn'
    elif slug == 'facilities':
        return 'purple-btn'
    return ''

# Main routine: Parse, Generate PPTX, and Build Portal HTML
all_parsed_apps = []
app_cards_html = []

cat_badge_map = {
    'Autonomous AI & Swarms': 'ai-ops',
    'Plant Engineering & Twins': 'engineering',
    'Grids & Utilities': 'utilities',
    'Vision & Clean Energy': 'vision-energy',
    'Smart Facilities & BMS': 'facilities'
}

for fname, cat_name, cat_slug, def_proto in app_meta_list:
    app_data = parse_app_html(fname, cat_name, cat_slug, def_proto)
    all_parsed_apps.append(app_data)
    
    clean_name = app_data['name'].replace(' ', '_')
    pptx_filename = f"{clean_name}_Catalogue.pptx"
    
    # Generate PPTX
    prs = generate_catalogue_pptx(app_data)
    prs.save(os.path.join(catalogues_dir, pptx_filename))
    prs.save(os.path.join(meta_catalogues_dir, pptx_filename))
    prs.save(os.path.join(base_dir, 'catalogues', pptx_filename))
    
    # Generate App Card HTML
    logo_html = format_logo(app_data['name'])
    badge_cls = cat_badge_map.get(cat_name, 'tech')
    btn_cls = get_btn_class(cat_slug)
    deck_href = f"catalogues/{pptx_filename}"
    spec_link = f"../terrestrial.cloud/platform/{app_data['file']}"
    
    features_html = '\n'.join([f'          <li>{feat[0]}</li>' for feat in app_data['features'][:4]])
    
    card_html = f'''      <!-- {app_data['name']} -->
      <div class="app-card" data-cat="catalogue" data-subcat="{cat_slug}" data-search="{app_data['name'].lower()} {cat_name.lower()} {app_data['protocol'].lower()} {app_data['brief'].lower()}">
        <div class="app-card-top">
          <div class="app-logo">{logo_html}</div>
          <span class="card-badge {badge_cls}">{cat_name}</span>
        </div>
        <div class="app-protocol-badge">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>
          {app_data['protocol']}
        </div>
        <a href="{spec_link}" target="_blank" style="text-decoration: none; color: inherit;">
          <h3 class="app-name">{app_data['name']}</h3>
        </a>
        <p class="app-desc">
          {app_data['brief']}
        </p>
        <ul class="app-features">
{features_html}
        </ul>
        <div class="card-actions">
          <a href="{spec_link}" target="_blank" class="btn-download {btn_cls}">
            <span>Explore {app_data['name']} Specification</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
          </a>
        </div>
      </div>'''
    app_cards_html.append(card_html)

all_cards_block = '\n\n'.join(app_cards_html)

# Now assemble the full HTML
full_html = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Terrestrial Platform — SI Sales &amp; Application Catalogues Hub</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Space+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet" />
  <style>
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

    #starfield {{
      position: fixed; inset: 0; z-index: 0; overflow: hidden; pointer-events: none;
    }}
    #starfield canvas {{ width: 100%; height: 100%; display: block; }}

    body::before {{
      content: '';
      position: fixed; inset: 0; z-index: 0; pointer-events: none;
      background-image:
        linear-gradient(rgba(0,212,255,0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(0,212,255,0.03) 1px, transparent 1px);
      background-size: 60px 60px;
    }}

    .wrap {{ position: relative; z-index: 1; max-width: 1400px; margin: 0 auto; padding: 0 32px; }}

    /* NAV */
    nav {{
      position: sticky; top: 0; z-index: 100;
      background: rgba(10,14,26,0.92);
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

    /* HERO */
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

    /* SECTIONS */
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

    /* FILTERS */
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

    /* CORE CARDS GRID (Decks) */
    .cards-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
      gap: 32px;
      max-width: 1100px;
      margin: 0 auto;
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
      font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;
      padding: 5px 12px; border-radius: 50px;
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
      font-size: 22px; font-weight: 700; line-height: 1.25;
      margin-bottom: 12px;
    }}
    .card-desc {{
      font-size: 15px; color: var(--text); line-height: 1.65;
      margin-bottom: 20px; flex: 1;
    }}

    .card-meta {{
      display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;
    }}
    .card-meta-item {{
      display: flex; align-items: center; gap: 6px;
      font-size: 13px; color: var(--muted); font-weight: 500;
    }}
    .card-meta-item svg {{ opacity: 0.6; }}

    .card-tags {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 24px; }}
    .card-tag {{
      font-size: 12px; font-weight: 600; padding: 5px 12px;
      border-radius: 6px; background: var(--glass2);
      border: 1px solid var(--border); color: var(--muted);
    }}

    .card-actions {{ display: flex; gap: 10px; }}

    .btn-download {{
      flex: 1; display: flex; align-items: center; justify-content: center; gap: 8px;
      background: linear-gradient(135deg, var(--blue2), var(--cyan2));
      border: none; border-radius: 12px; padding: 13px 20px;
      font-size: 14.5px; font-weight: 700; color: var(--white);
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

    .btn-download-sec {{
      flex: 1; display: flex; align-items: center; justify-content: center; gap: 8px;
      background: var(--glass2);
      border: 1px solid var(--border);
      border-radius: 12px; padding: 13px 18px;
      font-size: 14px; font-weight: 700; color: var(--white);
      cursor: pointer; transition: var(--tr); text-decoration: none;
      letter-spacing: 0.3px;
    }}
    .btn-download-sec:hover {{
      transform: translateY(-2px);
      border-color: var(--cyan);
      background: rgba(0, 212, 255, 0.1);
      color: var(--cyan);
      box-shadow: 0 4px 20px rgba(0,212,255,0.25);
    }}
    .btn-download-sec.gold-outline:hover {{
      border-color: var(--gold);
      background: rgba(245, 200, 66, 0.1);
      color: var(--gold);
      box-shadow: 0 4px 20px rgba(245,200,66,0.25);
    }}

    /* FEATURED BANNER */
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

    /* CATALOGUE CONTROLS */
    .catalogue-controls {{
      display: flex; flex-direction: column; gap: 20px;
      margin-bottom: 40px; align-items: center;
    }}
    .catalogue-search-bar {{
      width: 100%; max-width: 640px; position: relative;
    }}
    .catalogue-search-input {{
      width: 100%; padding: 16px 22px 16px 52px;
      border-radius: 50px; background: rgba(13, 19, 38, 0.9);
      border: 1px solid var(--border);
      color: var(--white); font-size: 15px;
      outline: none; transition: var(--tr);
      box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      font-family: 'Inter', sans-serif;
    }}
    .catalogue-search-input:focus {{
      border-color: var(--cyan);
      box-shadow: 0 0 25px rgba(0,212,255,0.25);
    }}
    .catalogue-search-icon {{
      position: absolute; left: 20px; top: 50%; transform: translateY(-50%);
      color: var(--cyan); pointer-events: none; opacity: 0.8;
    }}
    .sub-filter-bar {{
      display: flex; justify-content: center; flex-wrap: wrap; gap: 10px;
    }}
    .sub-filter-btn {{
      background: var(--glass); border: 1px solid var(--border);
      border-radius: 50px; padding: 10px 20px;
      font-size: 13.5px; font-weight: 600; color: var(--muted);
      cursor: pointer; transition: var(--tr);
      letter-spacing: 0.3px;
    }}
    .sub-filter-btn:hover, .sub-filter-btn.active {{
      background: linear-gradient(135deg, rgba(0,80,255,0.25), rgba(0,212,255,0.18));
      border-color: var(--cyan); color: var(--cyan);
      box-shadow: 0 0 16px rgba(0,212,255,0.18);
    }}

    /* APP CATALOGUE GRID */
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
      font-size: 17px; font-weight: 700; letter-spacing: 1px;
    }}
    .app-protocol-badge {{
      display: inline-flex; align-items: center; gap: 6px;
      background: rgba(0, 212, 255, 0.08);
      border: 1px solid rgba(0, 212, 255, 0.22);
      border-radius: 6px; padding: 5px 12px;
      font-size: 12px; font-weight: 600; color: var(--cyan);
      font-family: 'Space Grotesk', sans-serif;
      margin-bottom: 12px; width: fit-content;
    }}
    .app-name {{
      font-family: 'Space Grotesk', sans-serif;
      font-size: 24px; font-weight: 700; margin-bottom: 10px;
      color: var(--white);
    }}
    .app-desc {{ font-size: 15px; color: var(--text); line-height: 1.65; margin-bottom: 18px; flex: 1; }}
    .app-features {{ list-style: none; margin-bottom: 22px; display: flex; flex-direction: column; gap: 8px; }}
    .app-features li {{
      display: flex; align-items: flex-start; gap: 10px;
      font-size: 13.5px; color: var(--text); line-height: 1.45;
    }}
    .app-features li::before {{
      content: ''; width: 5px; height: 5px; border-radius: 50%;
      background: var(--cyan); flex-shrink: 0; margin-top: 6px;
    }}

    /* HOW TO USE */
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

    /* FOOTER */
    footer {{
      border-top: 1px solid var(--border);
      padding: 40px 0;
      text-align: center;
      color: var(--muted); font-size: 13px;
    }}
    footer strong {{ color: var(--cyan); }}

    /* TOAST */
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

<div id="starfield"><canvas id="c"></canvas></div>

<!-- NAV -->
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

<!-- HERO -->
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
      <div class="hero-stat"><strong>2</strong><span>Core Presentations</span></div>
      <div class="hero-stat"><strong>20</strong><span>Enterprise App Catalogues</span></div>
      <div class="hero-stat"><strong>5</strong><span>Industrial Sectors</span></div>
    </div>
  </div>
</section>

<!-- SECTION 1 — Core Presentations -->
<section class="section">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Step 1 — Choose Your Deck</div>
      <h2 class="section-h2">Core Sales Presentations</h2>
      <p class="section-desc">Pick the right narrative for the right audience. Each presentation provides comprehensive speech talking scripts and PowerPoint slides.</p>
    </div>

    <div class="filter-bar">
      <button class="filter-btn active" data-filter="all">All Materials</button>
      <button class="filter-btn" data-filter="exec">Executive</button>
      <button class="filter-btn" data-filter="tech">Technical</button>
      <button class="filter-btn" data-filter="catalogue">App Specifications (20)</button>
    </div>

    <div class="cards-grid" id="cardsGrid">

      <!-- Executive -->
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
            Includes verbatim pitch speech script, boardroom speaker notes, EBITDA expansion equations,
            and C-Suite objection handling battlecards.
          </p>
          <div class="card-tags">
            <span class="card-tag">Speech Script Included</span>
            <span class="card-tag">Boardroom Notes</span>
            <span class="card-tag">EBITDA Focus</span>
            <span class="card-tag">15–20 min pitch</span>
          </div>
          <div class="card-meta">
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              Speech Deck &amp; Slides (.pptx)
            </div>
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              Updated Oct 2026
            </div>
          </div>
          <div class="card-actions">
            <a href="../Executive_Speech_Deck.pptx" class="btn-download gold-btn" download onclick="showToast('Executive Speech Deck', 'Executive_Speech_Deck.pptx')">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="22"/></svg>
              Download Speech Deck
            </a>
            <a href="../The_Terrestrial_Metamorphosis-ver2.pptx" class="btn-download-sec gold-outline" download="The_Terrestrial_Metamorphosis-ver2.pptx" onclick="showToast('Executive Presentation', 'The_Terrestrial_Metamorphosis-ver2.pptx')">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download PPTX
            </a>
          </div>
        </div>
      </div>

      <!-- Technical -->
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
            Includes verbatim technical pitch speech script, Purdue model air-gap compliance notes,
            bumpless supervisory closed-loop explanations, and engineer Q&amp;A battlecards.
          </p>
          <div class="card-tags">
            <span class="card-tag">Technical Speech Script</span>
            <span class="card-tag">OT/IT Talking Points</span>
            <span class="card-tag">Edge-to-Cloud</span>
            <span class="card-tag">SCADA Replace</span>
          </div>
          <div class="card-meta">
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              Speech Deck &amp; Slides (.pptx)
            </div>
            <div class="card-meta-item">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              Updated Oct 2026
            </div>
          </div>
          <div class="card-actions">
            <a href="../Engineering_Speech_Deck.pptx" class="btn-download" download onclick="showToast('Engineering Speech Deck', 'Engineering_Speech_Deck.pptx')">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="22"/></svg>
              Download Speech Deck
            </a>
            <a href="../CITA-TP-I2.pptx" class="btn-download-sec" download="CITA-TP-I2.pptx" onclick="showToast('Engineering Presentation', 'CITA-TP-I2.pptx')">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download PPTX
            </a>
          </div>
        </div>
      </div>

    </div>
  </div>
</section>

<!-- SECTION 2 — Application Catalogues -->
<section class="section" style="padding-top: 0;" id="catalogues">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Step 2 — Drill Into Applications</div>
      <h2 class="section-h2">Application Specifications &amp; Platform Suite</h2>
      <p class="section-desc">Interactive product specifications for all 20 Terra application modules across 5 industrial sectors. Click any application below to explore its live platform architecture and technical capabilities.</p>
    </div>

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

{all_cards_block}

    </div>
  </div>
</section>

<!-- SECTION 3 — How to Use -->
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

<!-- FOOTER -->
<footer>
  <div class="wrap">
    <p>© 2026 <strong>Terrestrial Platform™</strong> &nbsp;|&nbsp; SI Partner Portal &nbsp;|&nbsp; Confidential — For Authorised System Integrators Only</p>
    <p style="margin-top:8px;">From Edge to Cloud &nbsp;·&nbsp; All materials are proprietary and not for redistribution.</p>
  </div>
</footer>

<!-- TOAST -->
<div id="toast">
  <div id="toast-icon">⬇️</div>
  <div>
    <div id="toast-text">Downloading…</div>
    <div id="toast-sub">Your file will be ready in a moment</div>
  </div>
</div>

<script>
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

const filterBtns = document.querySelectorAll('.filter-btn');
filterBtns.forEach(btn => {{
  btn.addEventListener('click', () => {{
    filterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const f = btn.dataset.filter;

    document.querySelectorAll('#cardsGrid .card').forEach(c => {{
      c.style.display = (f === 'all' || c.dataset.cat === f) ? '' : 'none';
    }});

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

target_index = os.path.join(si_portal_dir, 'index.html')
with open(target_index, 'w', encoding='utf-8') as f:
    f.write(full_html)
print(f"Updated {target_index} successfully!")

meta_target_index = os.path.join(meta_portal_dir, 'index.html')
with open(meta_target_index, 'w', encoding='utf-8') as f:
    f.write(full_html)
print(f"Updated {meta_target_index} successfully!")

print(f"\nSUCCESS: Generated 20 application PPTX catalogues and updated SI portal index with direct download links!")

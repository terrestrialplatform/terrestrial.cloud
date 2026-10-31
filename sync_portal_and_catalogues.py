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
            'AI': '#4900f3',
            'COGNIZANCE': '#7c3aed',
            'EMBODIMENT': '#0284c7',
            'HIVE': '#4900f3',
            'CC': '#7c3aed',
            'SCADA': '#0284c7',
            'AVATAR': '#0ea5e9',
            'ASSETS': '#4900f3',
            'PFD': '#7c3aed',
            'ALARM': '#e11d48',
            'GRIDX': '#059669',
            'GRIDW': '#0284c7',
            'GRIDG': '#d97706',
            'CORROGUARD': '#e11d48',
            'SOLAR': '#d97706',
            'EMS': '#059669',
            '3RDEYE': '#0284c7',
            'BMS': '#7c3aed',
            'HVAC': '#059669'
        }
        col = colors.get(suffix, '#4900f3')
        return f'<span style="color:#111827;font-weight:800;">TERRA</span><span style="color:{col};font-weight:900;">{suffix}</span>'
    elif u == 'TPIM':
        return '<span style="color:#4900f3;font-weight:900;">TPIM</span> <span style="color:#6b7280;font-size:11px;font-weight:600;">SCHEMA</span>'
    return f'<span style="color:#111827;font-weight:800;">{u}</span>'

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
    spec_link = f"../platform/{app_data['file']}"
    
    features_html = '\n'.join([f'          <li>{feat[0]}</li>' for feat in app_data['features'][:4]])
    
    card_html = f'''      <!-- {app_data['name']} -->
      <div class="app-card" data-cat="catalogue" data-subcat="{cat_slug}" data-search="{app_data['name'].lower()} {cat_name.lower()} {app_data['protocol'].lower()} {app_data['brief'].lower()}">
        <div class="app-card-top">
          <div class="app-logo">{logo_html}</div>
          <span class="card-badge {badge_cls}">{cat_name}</span>
        </div>
        <div class="app-protocol-badge">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>
          <span>{app_data['protocol']}</span>
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
        <div class="app-card-actions">
          <a href="{spec_link}" target="_blank" class="btn-explore">
            <span>Explore {app_data['name']} Specification</span>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
          </a>
        </div>
      </div>'''
    app_cards_html.append(card_html)

all_cards_block = '\n\n'.join(app_cards_html)

# Assemble Full HTML with LandingPage Theme & Header
full_html = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Terrestrial Platform — System Integrator Partner Portal</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet" />
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    :root {{
      --primary: #4900f3;
      --primary-hover: #3a00c7;
      --primary-light: #faf8ff;
      --primary-tint: #ede9fe;
      --primary-border: #e9d5ff;

      --secondary: #0ea5e9;
      --secondary-hover: #0284c7;
      --secondary-light: #f0f9ff;
      --secondary-tint: #e0f2fe;
      --secondary-border: #bae6fd;

      --gray-50: #f9fafb;
      --gray-100: #f3f4f6;
      --gray-200: #e5e7eb;
      --gray-300: #d1d5db;
      --gray-400: #9ca3af;
      --gray-500: #6b7280;
      --gray-600: #4b5563;
      --gray-700: #374151;
      --gray-800: #1f2937;
      --gray-900: #111827;

      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 18px;
      --radius-xl: 24px;
      --radius-full: 9999px;
      --tr: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    html {{ scroll-behavior: smooth; }}

    body {{
      font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background-color: #ffffff;
      color: var(--gray-900);
      overflow-x: hidden;
      min-height: 100vh;
    }}

    ::selection {{
      background: #ede9fe;
      color: #4900f3;
    }}

    .wrap {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 0 24px;
      position: relative;
      z-index: 2;
    }}

    /* HEADER / NAVBAR from LandingPage.tsx */
    .site-header {{
      position: sticky;
      top: 0;
      z-index: 100;
      background: rgba(255, 255, 255, 0.88);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--gray-100);
      transition: var(--tr);
    }}
    .header-inner {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 0 24px;
      height: 64px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }}
    .nav-brand {{
      display: flex;
      align-items: center;
      gap: 10px;
      text-decoration: none;
    }}
    .nav-logo {{
      height: 32px;
      width: auto;
      object-fit: contain;
      transition: transform 0.2s ease;
    }}
    .nav-logo:hover {{
      transform: scale(1.05);
    }}
    .nav-brand-text-block {{
      display: flex;
      flex-direction: column;
      line-height: 1.15;
    }}
    .nav-brand-title {{
      font-weight: 800;
      font-size: 13.5px;
      color: var(--gray-900);
      letter-spacing: 0.04em;
    }}
    .nav-brand-subtitle {{
      font-size: 8.5px;
      font-weight: 700;
      color: #0284c7;
      letter-spacing: 0.05em;
      text-transform: uppercase;
    }}
    .nav-menu {{
      display: flex;
      align-items: center;
      gap: 32px;
    }}
    .nav-link {{
      font-size: 14px;
      font-weight: 500;
      color: var(--gray-500);
      text-decoration: none;
      transition: color 0.2s ease;
    }}
    .nav-link:hover {{
      color: var(--gray-900);
    }}
    .nav-actions {{
      display: flex;
      align-items: center;
      gap: 12px;
    }}
    .text-center {{
      text-align: center;
    }}
    .btn-signin {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 16px;
      border-radius: var(--radius-sm);
      font-size: 14px;
      font-weight: 500;
      color: var(--gray-700);
      border: 1px solid var(--gray-200);
      background: #ffffff;
      text-decoration: none;
      transition: var(--tr);
      cursor: pointer;
    }}
    .btn-signin:hover {{
      border-color: var(--gray-300);
      background: var(--gray-50);
      color: var(--gray-900);
    }}
    .btn-apply-now {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 18px;
      border-radius: var(--radius-sm);
      font-size: 14px;
      font-weight: 500;
      color: #ffffff;
      background: #4900f3;
      text-decoration: none;
      box-shadow: 0 1px 3px rgba(73, 0, 243, 0.25);
      transition: var(--tr);
      cursor: pointer;
    }}
    .btn-apply-now:hover {{
      background: #3a00c7;
      box-shadow: 0 4px 14px rgba(73, 0, 243, 0.35);
      transform: translateY(-2px);
      color: #ffffff;
    }}
    .btn-apply-now .apply-arrow {{
      transition: transform 0.2s ease;
    }}
    .btn-apply-now:hover .apply-arrow {{
      transform: translateX(3px);
    }}

    .hero-cta-wrap {{
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 14px;
      margin-top: 36px;
      flex-wrap: wrap;
    }}
    .hero-cta-btn {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      padding: 14px 28px;
      border-radius: var(--radius-md);
      font-weight: 600;
      font-size: 15px;
      color: #ffffff;
      background: linear-gradient(135deg, #4900f3 0%, #3a00c7 100%);
      box-shadow: 0 10px 28px rgba(73, 0, 243, 0.28);
      text-decoration: none;
      transition: var(--tr);
      cursor: pointer;
    }}
    .hero-cta-btn:hover {{
      transform: translateY(-2px);
      box-shadow: 0 16px 36px rgba(73, 0, 243, 0.4);
      color: #ffffff;
    }}
    .hero-cta-btn .cta-arrow {{
      transition: transform 0.2s ease;
    }}
    .hero-cta-btn:hover .cta-arrow {{
      transform: translateX(4px);
    }}
    .hero-secondary-btn {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      padding: 14px 24px;
      border-radius: var(--radius-md);
      font-weight: 600;
      font-size: 15px;
      color: var(--gray-700);
      background: #ffffff;
      border: 1.5px solid var(--gray-200);
      text-decoration: none;
      transition: var(--tr);
      box-shadow: 0 2px 6px rgba(0,0,0,0.04);
      cursor: pointer;
    }}
    .hero-secondary-btn:hover {{
      border-color: var(--primary);
      color: var(--primary);
      background: var(--primary-light);
      transform: translateY(-2px);
    }}
    .live-dot-wrap {{
      position: relative;
      display: flex;
      height: 8px;
      width: 8px;
    }}
    .live-dot-ping {{
      position: absolute;
      inset: 0;
      border-radius: 50%;
      background-color: var(--primary);
      opacity: 0.75;
      animation: ping 1.5s cubic-bezier(0, 0, 0.2, 1) infinite;
    }}
    .live-dot {{
      position: relative;
      border-radius: 50%;
      height: 8px;
      width: 8px;
      background-color: var(--primary);
    }}
    @keyframes ping {{
      75%, 100% {{
        transform: scale(2);
        opacity: 0;
      }}
    }}

    /* HERO SECTION with LandingPage Pattern & Ambient Orbs */
    .hero-section {{
      position: relative;
      overflow: hidden;
      padding: 96px 0 90px;
      background: #ffffff;
      text-align: center;
    }}
    .bg-grid-pattern {{
      position: absolute;
      inset: 0;
      pointer-events: none;
      opacity: 0.55;
      background-image: radial-gradient(circle at 1px 1px, #d1d5db 1px, transparent 0);
      background-size: 36px 36px;
      mask-image: linear-gradient(to bottom, black 65%, transparent 100%);
      -webkit-mask-image: linear-gradient(to bottom, black 65%, transparent 100%);
    }}
    .ambient-glow-1 {{
      position: absolute;
      top: 30px;
      left: 50%;
      transform: translateX(-50%);
      width: 800px;
      height: 480px;
      border-radius: 50%;
      background: radial-gradient(ellipse, rgba(73, 0, 243, 0.12) 0%, rgba(124, 58, 237, 0.05) 40%, transparent 70%);
      filter: blur(60px);
      pointer-events: none;
    }}
    .ambient-glow-2 {{
      position: absolute;
      top: 120px;
      right: 5%;
      width: 400px;
      height: 400px;
      border-radius: 50%;
      background: radial-gradient(ellipse, rgba(14, 165, 233, 0.08) 0%, transparent 70%);
      filter: blur(60px);
      pointer-events: none;
    }}

    .hero-badge {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 18px;
      border-radius: var(--radius-full);
      font-size: 12.5px;
      font-weight: 600;
      border: 1px solid var(--primary-border);
      background: var(--primary-light);
      color: var(--primary);
      box-shadow: 0 1px 3px rgba(73, 0, 243, 0.06);
      margin: 0 auto 28px;
    }}
    .hero-title {{
      font-size: clamp(38px, 5.5vw, 64px);
      font-weight: 800;
      line-height: 1.12;
      letter-spacing: -0.03em;
      color: var(--gray-900);
      max-width: 900px;
      margin: 0 auto 24px;
      text-align: center;
    }}
    .hero-gradient-text {{
      background: linear-gradient(135deg, #4900f3 0%, #6320ee 50%, #7c3aed 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }}
    .hero-subtext {{
      font-size: clamp(16px, 1.9vw, 19px);
      color: var(--gray-600);
      max-width: 740px;
      margin: 0 auto;
      line-height: 1.7;
      font-weight: 400;
      text-align: center;
    }}

    /* STATS CARDS */
    .hero-stats-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 20px;
      max-width: 840px;
      margin: 0 auto;
    }}
    .stat-card {{
      position: relative;
      background: rgba(255, 255, 255, 0.85);
      backdrop-filter: blur(8px);
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-lg);
      padding: 24px 20px;
      box-shadow: 0 4px 20px -4px rgba(0, 0, 0, 0.05);
      transition: var(--tr);
      text-align: center;
    }}
    .stat-card:hover {{
      transform: translateY(-3px);
      border-color: var(--primary-border);
      box-shadow: 0 12px 32px -6px rgba(73, 0, 243, 0.12);
    }}
    .stat-accent {{
      position: absolute;
      top: 0;
      left: 50%;
      transform: translateX(-50%);
      width: 48px;
      height: 3px;
      border-radius: 0 0 9999px 9999px;
      background: linear-gradient(90deg, transparent, var(--primary), transparent);
    }}
    .stat-num {{
      font-size: 38px;
      font-weight: 800;
      color: var(--gray-900);
      line-height: 1;
      letter-spacing: -0.02em;
      margin-bottom: 8px;
    }}
    .stat-label {{
      font-size: 13px;
      font-weight: 500;
      color: var(--gray-500);
    }}

    /* SECTION BASE */
    .section {{
      padding: 72px 0;
      position: relative;
    }}
    .section-header {{
      text-align: center;
      margin-bottom: 48px;
    }}
    .section-tag {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 14px;
      border-radius: var(--radius-full);
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      color: var(--gray-500);
      background: var(--gray-100);
      margin-bottom: 14px;
    }}
    .section-h2 {{
      font-size: clamp(28px, 3.8vw, 42px);
      font-weight: 800;
      letter-spacing: -0.02em;
      color: var(--gray-900);
      margin-bottom: 12px;
    }}
    .section-desc {{
      color: var(--gray-500);
      font-size: 16px;
      max-width: 640px;
      margin: 0 auto;
      line-height: 1.65;
    }}

    /* FILTER BAR */
    .filter-bar {{
      display: flex;
      justify-content: center;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 44px;
    }}
    .filter-btn {{
      background: #ffffff;
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-full);
      padding: 9px 22px;
      font-size: 13.5px;
      font-weight: 500;
      color: var(--gray-600);
      cursor: pointer;
      transition: var(--tr);
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }}
    .filter-btn:hover {{
      border-color: var(--primary-border);
      color: var(--primary);
      background: var(--primary-light);
    }}
    .filter-btn.active {{
      background: var(--primary);
      color: #ffffff;
      border-color: var(--primary);
      box-shadow: 0 4px 14px rgba(73, 0, 243, 0.28);
      font-weight: 600;
    }}

    /* CORE CARDS GRID (Decks) */
    .cards-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(440px, 1fr));
      gap: 28px;
      max-width: 1050px;
      margin: 0 auto;
    }}
    .deck-card {{
      position: relative;
      border-radius: var(--radius-xl);
      padding: 38px 34px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      transition: var(--tr);
    }}
    .deck-card:hover {{
      transform: translateY(-4px);
      box-shadow: 0 20px 45px -10px rgba(0, 0, 0, 0.1);
    }}
    .deck-card-exec {{
      border: 2px solid #4900f3;
      background: #faf9ff;
      box-shadow: 0 4px 25px -4px rgba(73, 0, 243, 0.08);
    }}
    .deck-card-tech {{
      border: 2px solid #0ea5e9;
      background: #f0f9ff;
      box-shadow: 0 4px 25px -4px rgba(14, 165, 233, 0.08);
    }}
    .deck-card-top {{
      display: flex;
      align-items: center;
      gap: 14px;
      margin-bottom: 20px;
    }}
    .deck-icon-wrap {{
      width: 48px;
      height: 48px;
      border-radius: var(--radius-md);
      display: flex;
      align-items: center;
      justify-content: center;
      color: #ffffff;
      flex-shrink: 0;
    }}
    .exec-icon-wrap {{
      background: #4900f3;
      box-shadow: 0 6px 16px rgba(73, 0, 243, 0.3);
    }}
    .tech-icon-wrap {{
      background: #0ea5e9;
      box-shadow: 0 6px 16px rgba(14, 165, 233, 0.3);
    }}
    .deck-sub-badge {{
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-bottom: 2px;
    }}
    .exec-sub {{ color: #4900f3; }}
    .tech-sub {{ color: #0284c7; }}
    .deck-cat-badge {{
      display: inline-block;
      padding: 2px 10px;
      border-radius: var(--radius-full);
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
    }}
    .exec-badge {{
      background: #ede9fe;
      color: #4900f3;
      border: 1px solid #ddd6fe;
    }}
    .tech-badge {{
      background: #e0f2fe;
      color: #0284c7;
      border: 1px solid #bae6fd;
    }}
    .deck-title {{
      font-size: 24px;
      font-weight: 700;
      color: var(--gray-900);
      margin-bottom: 12px;
      line-height: 1.25;
      letter-spacing: -0.01em;
    }}
    .deck-desc {{
      font-size: 14.5px;
      color: var(--gray-600);
      line-height: 1.65;
      margin-bottom: 22px;
      flex: 1;
    }}
    .deck-tags {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 22px;
    }}
    .deck-tag {{
      font-size: 12px;
      font-weight: 600;
      padding: 5px 12px;
      border-radius: 6px;
    }}
    .exec-tag {{
      background: #ede9fe;
      color: #4900f3;
      border: 1px solid #ddd6fe;
    }}
    .tech-tag {{
      background: #e0f2fe;
      color: #0284c7;
      border: 1px solid #bae6fd;
    }}
    .deck-meta {{
      display: flex;
      gap: 18px;
      margin-bottom: 24px;
      flex-wrap: wrap;
    }}
    .deck-meta-item {{
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12.5px;
      color: var(--gray-500);
      font-weight: 500;
    }}
    .deck-actions {{
      display: flex;
      gap: 12px;
    }}
    .btn-deck-primary {{
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      border: none;
      border-radius: var(--radius-md);
      padding: 13px 20px;
      font-size: 14px;
      font-weight: 600;
      color: #ffffff;
      cursor: pointer;
      transition: var(--tr);
      text-decoration: none;
    }}
    .exec-btn {{
      background: linear-gradient(135deg, #4900f3 0%, #3a00c7 100%);
      box-shadow: 0 6px 20px rgba(73, 0, 243, 0.28);
    }}
    .exec-btn:hover {{
      box-shadow: 0 10px 28px rgba(73, 0, 243, 0.4);
      transform: translateY(-2px);
    }}
    .tech-btn {{
      background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%);
      box-shadow: 0 6px 20px rgba(14, 165, 233, 0.28);
    }}
    .tech-btn:hover {{
      box-shadow: 0 10px 28px rgba(14, 165, 233, 0.4);
      transform: translateY(-2px);
    }}
    .btn-deck-secondary {{
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      background: #ffffff;
      border-radius: var(--radius-md);
      padding: 13px 20px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: var(--tr);
      text-decoration: none;
    }}
    .exec-outline {{
      border: 1.5px solid #4900f3;
      color: #4900f3;
    }}
    .exec-outline:hover {{
      background: #ede9fe;
      transform: translateY(-2px);
    }}
    .tech-outline {{
      border: 1.5px solid #0ea5e9;
      color: #0ea5e9;
    }}
    .tech-outline:hover {{
      background: #e0f2fe;
      transform: translateY(-2px);
    }}

    /* FEATURED BANNER from LandingPage Bottom CTA Style */
    .featured-banner {{
      background: linear-gradient(135deg, #4900f3 0%, #3500b8 100%);
      border-radius: var(--radius-xl);
      padding: 44px 48px;
      display: flex;
      align-items: center;
      gap: 40px;
      margin-bottom: 44px;
      position: relative;
      overflow: hidden;
      box-shadow: 0 16px 40px -8px rgba(73, 0, 243, 0.35);
      color: #ffffff;
    }}
    .featured-banner::before {{
      content: '';
      position: absolute;
      top: 0;
      right: 0;
      width: 380px;
      height: 380px;
      border-radius: 50%;
      background: rgba(255, 255, 255, 0.1);
      filter: blur(50px);
      pointer-events: none;
      transform: translate(30%, -30%);
    }}
    .featured-banner-text {{
      flex: 1;
      position: relative;
      z-index: 1;
    }}
    .featured-banner-text h3 {{
      font-size: 26px;
      font-weight: 800;
      letter-spacing: -0.02em;
      margin-bottom: 10px;
      color: #ffffff;
    }}
    .featured-banner-text p {{
      font-size: 15px;
      color: rgba(255, 255, 255, 0.88);
      line-height: 1.65;
      max-width: 620px;
    }}
    .featured-banner-img {{
      width: 170px;
      height: auto;
      flex-shrink: 0;
      filter: drop-shadow(0 8px 24px rgba(0, 0, 0, 0.25));
      position: relative;
      z-index: 1;
    }}

    /* CATALOGUE CONTROLS */
    .catalogue-controls {{
      display: flex;
      flex-direction: column;
      gap: 18px;
      margin-bottom: 40px;
      align-items: center;
    }}
    .catalogue-search-bar {{
      width: 100%;
      max-width: 600px;
      position: relative;
    }}
    .catalogue-search-input {{
      width: 100%;
      padding: 15px 22px 15px 50px;
      border-radius: var(--radius-full);
      background: #ffffff;
      border: 1.5px solid var(--gray-200);
      color: var(--gray-900);
      font-size: 14.5px;
      outline: none;
      transition: var(--tr);
      box-shadow: 0 4px 16px rgba(0,0,0,0.03);
      font-family: inherit;
    }}
    .catalogue-search-input:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 0 4px rgba(73, 0, 243, 0.12);
    }}
    .catalogue-search-icon {{
      position: absolute;
      left: 18px;
      top: 50%;
      transform: translateY(-50%);
      color: var(--primary);
      pointer-events: none;
      opacity: 0.85;
    }}
    .sub-filter-bar {{
      display: flex;
      justify-content: center;
      flex-wrap: wrap;
      gap: 8px;
    }}
    .sub-filter-btn {{
      background: #ffffff;
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-full);
      padding: 8px 18px;
      font-size: 13px;
      font-weight: 500;
      color: var(--gray-600);
      cursor: pointer;
      transition: var(--tr);
    }}
    .sub-filter-btn:hover {{
      border-color: var(--primary-border);
      color: var(--primary);
      background: var(--primary-light);
    }}
    .sub-filter-btn.active {{
      background: var(--primary-tint);
      color: var(--primary);
      border-color: var(--primary-border);
      font-weight: 600;
      box-shadow: 0 2px 8px rgba(73, 0, 243, 0.1);
    }}

    /* APP CATALOGUE GRID (20 Apps) */
    .app-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
      gap: 24px;
    }}
    .app-card {{
      background: #ffffff;
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-lg);
      padding: 26px;
      transition: var(--tr);
      position: relative;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 0 2px 10px -2px rgba(0, 0, 0, 0.04);
    }}
    .app-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 3px;
      background: linear-gradient(90deg, transparent, var(--primary), transparent);
      opacity: 0;
      transition: var(--tr);
    }}
    .app-card:hover {{
      border-color: #c7d2fe;
      transform: translateY(-4px);
      box-shadow: 0 16px 36px -6px rgba(73, 0, 243, 0.12);
    }}
    .app-card:hover::before {{
      opacity: 1;
    }}

    .app-card-top {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 12px;
      gap: 10px;
    }}
    .app-logo {{
      font-size: 17px;
      font-weight: 700;
      letter-spacing: 0.5px;
    }}
    .card-badge {{
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      padding: 4px 10px;
      border-radius: var(--radius-full);
    }}
    .card-badge.ai-ops {{ background: #ede9fe; color: #4900f3; border: 1px solid #ddd6fe; }}
    .card-badge.engineering {{ background: #e0f2fe; color: #0284c7; border: 1px solid #bae6fd; }}
    .card-badge.utilities {{ background: #dcfce7; color: #16a34a; border: 1px solid #bbf7d0; }}
    .card-badge.vision-energy {{ background: #fef3c7; color: #d97706; border: 1px solid #fde68a; }}
    .card-badge.facilities {{ background: #f3e8ff; color: #9333ea; border: 1px solid #e9d5ff; }}

    .app-protocol-badge {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--gray-50);
      border: 1px solid var(--gray-200);
      border-radius: 6px;
      padding: 4px 10px;
      font-size: 11.5px;
      font-weight: 600;
      color: var(--gray-600);
      font-family: inherit;
      margin-bottom: 12px;
      width: fit-content;
    }}
    .app-name {{
      font-size: 22px;
      font-weight: 700;
      margin-bottom: 8px;
      color: var(--gray-900);
      letter-spacing: -0.01em;
      transition: color 0.2s ease;
    }}
    .app-name:hover {{
      color: var(--primary);
    }}
    .app-desc {{
      font-size: 14px;
      color: var(--gray-600);
      line-height: 1.6;
      margin-bottom: 18px;
      flex: 1;
    }}
    .app-features {{
      list-style: none;
      margin-bottom: 22px;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }}
    .app-features li {{
      display: flex;
      align-items: flex-start;
      gap: 9px;
      font-size: 13px;
      color: var(--gray-700);
      line-height: 1.45;
    }}
    .app-features li::before {{
      content: '';
      width: 5px;
      height: 5px;
      border-radius: 50%;
      background: var(--primary);
      flex-shrink: 0;
      margin-top: 6px;
    }}

    .app-card-actions {{
      margin-top: auto;
    }}
    .btn-explore {{
      width: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      background: var(--primary-light);
      border: 1.5px solid #e0e7ff;
      border-radius: var(--radius-md);
      padding: 12px 18px;
      font-size: 13.5px;
      font-weight: 600;
      color: var(--primary);
      cursor: pointer;
      transition: var(--tr);
      text-decoration: none;
    }}
    .btn-explore:hover {{
      background: var(--primary);
      border-color: var(--primary);
      color: #ffffff;
      box-shadow: 0 6px 20px rgba(73, 0, 243, 0.25);
      transform: translateY(-2px);
    }}

    /* HOW IT WORKS from LandingPage.tsx */
    .how-it-works-section {{
      background: var(--gray-50);
      border-top: 1px solid var(--gray-100);
    }}
    .steps-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 24px;
    }}
    .step-card {{
      background: #ffffff;
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-lg);
      padding: 30px 24px;
      text-align: center;
      transition: var(--tr);
      box-shadow: 0 1px 3px rgba(0,0,0,0.04);
      display: flex;
      flex-direction: column;
      align-items: center;
    }}
    .step-card:hover {{
      border-color: var(--primary-border);
      transform: translateY(-4px);
      box-shadow: 0 12px 30px -4px rgba(73, 0, 243, 0.12);
    }}
    .step-badge {{
      font-size: 11px;
      font-weight: 800;
      color: #818cf8;
      letter-spacing: 0.08em;
      margin-bottom: 14px;
      transition: color 0.2s ease;
    }}
    .step-card:hover .step-badge {{
      color: var(--primary);
    }}
    .step-icon-circle {{
      width: 48px;
      height: 48px;
      border-radius: var(--radius-md);
      background: var(--primary-tint);
      color: var(--primary);
      display: flex;
      align-items: center;
      justify-content: center;
      margin-bottom: 16px;
      transition: transform 0.3s ease;
    }}
    .step-card:hover .step-icon-circle {{
      transform: scale(1.1);
    }}
    .step-title {{
      font-size: 16px;
      font-weight: 700;
      color: var(--gray-900);
      margin-bottom: 8px;
    }}
    .step-desc {{
      font-size: 13px;
      color: var(--gray-500);
      line-height: 1.6;
    }}

    /* FOOTER from LandingPage.tsx */
    .site-footer {{
      background: #ffffff;
      border-top: 1px solid var(--gray-200);
      padding: 32px 0;
    }}
    .footer-inner {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 0 24px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }}
    @media (min-width: 768px) {{
      .footer-inner {{
        flex-direction: row;
      }}
    }}
    .footer-brand {{
      display: flex;
      align-items: center;
      gap: 10px;
      text-decoration: none;
    }}
    .footer-logo {{
      height: 24px;
      width: auto;
      object-fit: contain;
    }}
    .footer-brand-text {{
      font-size: 14px;
      font-weight: 500;
      color: var(--gray-700);
    }}
    .footer-copyright {{
      font-size: 12px;
      color: var(--gray-400);
    }}

    /* TOAST */
    #toast {{
      position: fixed;
      bottom: 28px;
      right: 28px;
      z-index: 9999;
      background: #ffffff;
      border: 1.5px solid var(--primary-border);
      border-radius: var(--radius-md);
      padding: 14px 20px;
      display: flex;
      align-items: center;
      gap: 14px;
      box-shadow: 0 10px 30px rgba(73, 0, 243, 0.16);
      transform: translateX(120%);
      transition: transform 0.35s cubic-bezier(0.4, 0, 0.2, 1);
      max-width: 380px;
    }}
    #toast.show {{
      transform: translateX(0);
    }}
    #toast-icon {{
      font-size: 20px;
    }}
    #toast-text {{
      font-size: 14px;
      font-weight: 600;
      color: var(--gray-900);
    }}
    #toast-sub {{
      font-size: 12px;
      color: var(--gray-500);
      margin-top: 2px;
    }}

    @media (max-width: 900px) {{
      .nav-menu {{ display: none; }}
    }}
    @media (max-width: 768px) {{
      .cards-grid {{ grid-template-columns: 1fr; }}
      .hero-stats-grid {{ grid-template-columns: 1fr; }}
      .featured-banner {{ flex-direction: column; text-align: center; padding: 32px 24px; }}
      .featured-banner-img {{ width: 140px; }}
      .app-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>

<!-- STICKY HEADER matching LandingPage.tsx -->
<header class="site-header">
  <div class="header-inner">
    <a class="nav-brand" href="index.html">
      <img class="nav-logo" src="../TP8-F.png" alt="TP Platform Logo" />
      <div class="nav-brand-text-block">
        <span class="nav-brand-title">TERRESTRIAL PLATFORM</span>
        <span class="nav-brand-subtitle">ENTERPRISE INDUSTRIAL INTELLIGENCE SUITE</span>
      </div>
    </a>

    <nav class="nav-menu">
      <a href="#core-decks" class="nav-link">Core Presentations</a>
      <a href="#catalogues" class="nav-link">App Specifications (20)</a>
    </nav>

    <div class="nav-actions">
      <a href="https://auth.innovativeengineers.co.in/realms/Terrestrial-Platform/protocol/openid-connect/auth?client_id=terra-si&redirect_uri=https%3A%2F%2Fchannelpartner.terrestrial.cloud%2Fauth%2Fcallback&response_type=code&scope=openid%20profile%20email&state=b53b52a426f3a7117205b81165dbddf6" target="_blank" class="btn-signin">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"></path>
          <polyline points="10 17 15 12 10 7"></polyline>
          <line x1="15" y1="12" x2="3" y2="12"></line>
        </svg>
        <span>Sign In</span>
      </a>
      <a href="https://channelpartner.terrestrial.cloud/apply" target="_blank" class="btn-apply-now">
        <span>Apply Now</span>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="apply-arrow">
          <line x1="5" y1="12" x2="19" y2="12"></line>
          <polyline points="12 5 19 12 12 19"></polyline>
        </svg>
      </a>
    </div>
  </div>
</header>

<!-- HERO SECTION with Light Pattern & Orbs -->
<section class="hero-section">
  <div class="bg-grid-pattern"></div>
  <div class="ambient-glow-1"></div>
  <div class="ambient-glow-2"></div>

  <div class="wrap text-center">
    <div class="hero-badge">
      <span class="live-dot-wrap">
        <span class="live-dot-ping"></span>
        <span class="live-dot"></span>
      </span>
      <span>System Integrator Partner Portal</span>
    </div>

    <h1 class="hero-title">
      Your Complete<br/>
      <span class="hero-gradient-text">Sales &amp; Application Toolkit</span>
    </h1>

    <p class="hero-subtext">
      Everything you need to pitch, present &amp; close deals with confidence.
      Download tailored presentations for C-Suite executives, OT/IT engineers,
      and drill down into 20 application-wise specifications for every industrial domain.
    </p>

    <div class="hero-cta-wrap">
      <a href="https://channelpartner.terrestrial.cloud/apply" target="_blank" class="hero-cta-btn">
        <span>Apply as Channel Partner</span>
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="cta-arrow">
          <line x1="5" y1="12" x2="19" y2="12"></line>
          <polyline points="12 5 19 12 12 19"></polyline>
        </svg>
      </a>
      <a href="#catalogues" class="hero-secondary-btn">
        <span>Explore Specifications</span>
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </a>
    </div>
  </div>
</section>

<!-- SECTION 1 — Core Presentations -->
<section class="section" id="core-decks" style="background: #ffffff; border-top: 1px solid var(--gray-100);">
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

      <!-- Executive Card -->
      <div class="deck-card deck-card-exec" data-cat="exec">
        <div class="deck-card-top">
          <div class="deck-icon-wrap exec-icon-wrap">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle></svg>
          </div>
          <div>
            <div class="deck-sub-badge exec-sub">For C-Suite &amp; Board</div>
            <div class="deck-cat-badge exec-badge">Executive</div>
          </div>
        </div>
        <h3 class="deck-title">Presentation for Executive Personnel</h3>
        <p class="deck-desc">
          A high-impact, business-value-first narrative designed for CEOs, CFOs, VPs and Plant Heads.
          Includes verbatim pitch speech script, boardroom speaker notes, EBITDA expansion equations,
          and C-Suite objection handling battlecards.
        </p>
        <div class="deck-tags">
          <span class="deck-tag exec-tag">Speech Script Included</span>
          <span class="deck-tag exec-tag">Boardroom Notes</span>
          <span class="deck-tag exec-tag">EBITDA Focus</span>
          <span class="deck-tag exec-tag">15–20 min pitch</span>
        </div>
        <div class="deck-meta">
          <div class="deck-meta-item">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            <span>PowerPoint (.pptx) &amp; Speech Notes</span>
          </div>
          <div class="deck-meta-item">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            <span>Updated Oct 2026</span>
          </div>
        </div>
        <div class="deck-actions">
          <a href="../Executive_Speech_Deck.pptx" class="btn-deck-primary exec-btn" download onclick="showToast('Executive Speech Deck', 'Executive_Speech_Deck.pptx')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="22"/></svg>
            <span>Download Speech Deck</span>
          </a>
          <a href="../The_Terrestrial_Metamorphosis-ver2.pptx" class="btn-deck-secondary exec-outline" download="The_Terrestrial_Metamorphosis-ver2.pptx" onclick="showToast('Executive Presentation', 'The_Terrestrial_Metamorphosis-ver2.pptx')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            <span>Download PPTX</span>
          </a>
        </div>
      </div>

      <!-- Technical Card -->
      <div class="deck-card deck-card-tech" data-cat="tech">
        <div class="deck-card-top">
          <div class="deck-icon-wrap tech-icon-wrap">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
          </div>
          <div>
            <div class="deck-sub-badge tech-sub">For OT/IT &amp; Automation</div>
            <div class="deck-cat-badge tech-badge">Technical</div>
          </div>
        </div>
        <h3 class="deck-title">Presentation for Automation &amp; Engineering Personnel</h3>
        <p class="deck-desc">
          A deep-dive technical deck for OT Engineers, Automation Leads, SCADA/DCS teams, and Plant Managers.
          Includes verbatim technical pitch speech script, Purdue model air-gap compliance notes,
          bumpless supervisory closed-loop explanations, and engineer Q&amp;A battlecards.
        </p>
        <div class="deck-tags">
          <span class="deck-tag tech-tag">Technical Speech Script</span>
          <span class="deck-tag tech-tag">OT/IT Talking Points</span>
          <span class="deck-tag tech-tag">Edge-to-Cloud</span>
          <span class="deck-tag tech-tag">SCADA Replace</span>
        </div>
        <div class="deck-meta">
          <div class="deck-meta-item">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            <span>PowerPoint (.pptx) &amp; Technical Notes</span>
          </div>
          <div class="deck-meta-item">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            <span>Updated Oct 2026</span>
          </div>
        </div>
        <div class="deck-actions">
          <a href="../Engineering_Speech_Deck.pptx" class="btn-deck-primary tech-btn" download onclick="showToast('Engineering Speech Deck', 'Engineering_Speech_Deck.pptx')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="22"/></svg>
            <span>Download Speech Deck</span>
          </a>
          <a href="../CITA-TP-I2.pptx" class="btn-deck-secondary tech-outline" download="CITA-TP-I2.pptx" onclick="showToast('Engineering Presentation', 'CITA-TP-I2.pptx')">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            <span>Download PPTX</span>
          </a>
        </div>
      </div>

    </div>
  </div>
</section>

<!-- SECTION 2 — Application Specifications -->
<section class="section" id="catalogues" style="background: var(--gray-50); border-top: 1px solid var(--gray-100);">
  <div class="wrap">
    <div class="section-header">
      <div class="section-tag">Step 2 — Drill Into Applications</div>
      <h2 class="section-h2">Application Specifications &amp; Platform Suite</h2>
      <p class="section-desc">Interactive product specifications for all 20 Terra application modules across 5 industrial sectors. Click any application below to explore its live platform architecture and technical capabilities.</p>
    </div>

    <!-- Featured Purple Banner -->
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
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input type="text" id="catalogueSearch" class="catalogue-search-input" placeholder="Search 20 applications by name, category, protocol, or capability..." />
      </div>

      <div class="sub-filter-bar">
        <button class="sub-filter-btn active" data-sub="all">All Applications (20)</button>
        <button class="sub-filter-btn" data-sub="ai-ops">Autonomous AI &amp; Swarms (5)</button>
        <button class="sub-filter-btn" data-sub="engineering">Plant Engineering &amp; Twins (6)</button>
        <button class="sub-filter-btn" data-sub="utilities">Grids &amp; Utilities (4)</button>
        <button class="sub-filter-btn" data-sub="vision-energy">Vision &amp; Clean Energy (3)</button>
        <button class="sub-filter-btn" data-sub="facilities">Smart Facilities &amp; BMS (2)</button>
      </div>
    </div>

    <!-- APP CARDS GRID -->
    <div class="app-grid" id="appGrid">
{all_cards_block}
    </div>

  </div>
</section>

<!-- SECTION 3: HOW IT WORKS from LandingPage.tsx -->
<section class="section how-it-works-section" id="how-it-works">
  <div class="wrap text-center">
    <div class="section-tag">Process</div>
    <h2 class="section-h2">How It Works</h2>
    <p class="section-desc">From partner preparation to closed industrial enterprise contracts in four simple steps.</p>

    <div class="steps-grid" style="margin-top: 40px;">
      <div class="step-card">
        <div class="step-badge">STEP 01</div>
        <div class="step-icon-circle">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M22 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
        </div>
        <h3 class="step-title">Choose Your Narrative</h3>
        <p class="step-desc">Select Executive Pitch for C-Suite boardrooms or Technical Deep-Dive for OT/IT engineering teams.</p>
      </div>

      <div class="step-card">
        <div class="step-badge">STEP 02</div>
        <div class="step-icon-circle">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="7"></circle><polyline points="8.21 13.89 7 23 12 20 17 23 15.79 13.88"></polyline></svg>
        </div>
        <h3 class="step-title">Study Speech Guide</h3>
        <p class="step-desc">Review verbatim word-for-word scripts, stage cues, and EBITDA value propositions before the meeting.</p>
      </div>

      <div class="step-card">
        <div class="step-badge">STEP 03</div>
        <div class="step-icon-circle">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"></path><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"></path></svg>
        </div>
        <h3 class="step-title">Drill Into App Specs</h3>
        <p class="step-desc">Explore tailored live specifications for all 20 modules to address customer-specific operational needs.</p>
      </div>

      <div class="step-card">
        <div class="step-badge">STEP 04</div>
        <div class="step-icon-circle">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>
        </div>
        <h3 class="step-title">Deploy &amp; Earn ARR</h3>
        <p class="step-desc">Initiate non-intrusive 48-hour shadow pilots with full co-selling and turnkey software subscription margins.</p>
      </div>
    </div>
  </div>
</section>

<!-- FOOTER from LandingPage.tsx -->
<footer class="site-footer">
  <div class="footer-inner">
    <a class="footer-brand" href="index.html">
      <img class="footer-logo" src="../TP8-F.png" alt="TP Platform Logo" />
      <div class="nav-brand-text-block">
        <span style="font-weight:800;font-size:12px;color:var(--gray-900);letter-spacing:0.04em;">TERRESTRIAL PLATFORM</span>
        <span style="font-size:8px;font-weight:700;color:#0284c7;letter-spacing:0.05em;text-transform:uppercase;">ENTERPRISE INDUSTRIAL INTELLIGENCE SUITE</span>
      </div>
    </a>
    <span class="footer-copyright">© 2026 Innovation Intellect. All rights reserved.</span>
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
// Filter Core Decks
const filterBtns = document.querySelectorAll('.filter-btn');
filterBtns.forEach(btn => {{
  btn.addEventListener('click', () => {{
    filterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const f = btn.dataset.filter;

    document.querySelectorAll('#cardsGrid .deck-card').forEach(c => {{
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

// Search & Sub-category Filter for 20 Applications
const searchInput = document.getElementById('catalogueSearch');
const subFilterBtns = document.querySelectorAll('.sub-filter-btn');
const appCards = document.querySelectorAll('#appGrid .app-card');

let currentSub = 'all';
let currentQuery = '';

function filterApps() {{
  appCards.forEach(card => {{
    const sub = card.dataset.subcat || '';
    const searchData = card.dataset.search || '';
    const matchSub = (currentSub === 'all' || sub === currentSub);
    const matchQuery = (!currentQuery || searchData.includes(currentQuery));

    if (matchSub && matchQuery) {{
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

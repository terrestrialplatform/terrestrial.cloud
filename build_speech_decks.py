import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Directories
website_dir = r"c:\Users\yashj\Downloads\website"
meta_dir = r"c:\Users\yashj\Downloads\Terrestrial_metamorphosis"

# Color Palette (Dark Cybernetic Navy Theme)
c_navy = RGBColor(10, 14, 26)        # #0a0e1a
c_card = RGBColor(18, 24, 46)        # #12182e
c_card_highlight = RGBColor(24, 32, 60)
c_border = RGBColor(0, 212, 255)     # #00d4ff
c_border_sub = RGBColor(30, 41, 59)  # #1e293b
c_cyan = RGBColor(0, 212, 255)       # #00d4ff
c_gold = RGBColor(245, 200, 66)      # #f5c842
c_white = RGBColor(255, 255, 255)    # #ffffff
c_text = RGBColor(200, 216, 240)     # #c8d8f0
c_muted = RGBColor(148, 163, 184)    # #94a3b8
c_green = RGBColor(0, 230, 142)      # #00e68e

def create_deck(title_deck, deck_tag, slides_data, output_paths):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for slide_info in slides_data:
        s = prs.slides.add_slide(blank_layout)
        
        # Background
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = c_navy
        bg.line.fill.background()

        # Top Cyan Accent Line
        top_bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = c_cyan
        top_bar.line.fill.background()

        # Header
        tb_head = s.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
        tf_head = tb_head.text_frame
        tf_head.word_wrap = True
        
        p0 = tf_head.paragraphs[0]
        p0.text = f"{deck_tag}  |  {slide_info.get('step_tag', 'EXECUTIVE PITCH GUIDE')}"
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = c_cyan
        p0.space_after = Pt(4)

        p1 = tf_head.add_paragraph()
        p1.text = slide_info['title']
        p1.font.size = Pt(25)
        p1.font.bold = True
        p1.font.color.rgb = c_white

        # Slide Layout: Left Box (Speech Talking Script) & Right Box (Key Points & Takeaways)
        left_w = Inches(7.2)
        right_w = Inches(4.2)
        top_pos = Inches(1.7)
        height = Inches(4.8)
        gap = Inches(0.3)

        # LEFT BOX: Speech Script
        left_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top_pos, left_w, height)
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = c_card
        left_box.line.color.rgb = c_border
        
        ltf = left_box.text_frame
        ltf.word_wrap = True
        
        lp0 = ltf.paragraphs[0]
        lp0.text = "VERBATIM PRESENTATION SPEECH SCRIPT (WHAT TO SAY)"
        lp0.font.size = Pt(11)
        lp0.font.bold = True
        lp0.font.color.rgb = c_gold
        lp0.space_after = Pt(10)

        for para in slide_info['speech_script']:
            lp = ltf.add_paragraph()
            lp.text = para
            lp.font.size = Pt(13)
            lp.font.color.rgb = c_white
            lp.space_after = Pt(10)

        # RIGHT BOX: Key Talking Points & Delivery Notes
        right_x = Inches(0.8) + left_w + gap
        right_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, top_pos, right_w, height)
        right_box.fill.solid()
        right_box.fill.fore_color.rgb = c_card_highlight
        right_box.line.color.rgb = c_border_sub
        
        rtf = right_box.text_frame
        rtf.word_wrap = True
        
        rp0 = rtf.paragraphs[0]
        rp0.text = "SPEAKER CUES & VALUE DRIVERS"
        rp0.font.size = Pt(11)
        rp0.font.bold = True
        rp0.font.color.rgb = c_cyan
        rp0.space_after = Pt(10)

        for pt in slide_info['talking_points']:
            rp = rtf.add_paragraph()
            rp.text = f"- {pt}"
            rp.font.size = Pt(12.5)
            rp.font.color.rgb = c_text
            rp.space_after = Pt(8)

        # Tip / Objection Cues
        if 'tip' in slide_info:
            rtip = rtf.add_paragraph()
            rtip.text = f"PRESENTER TIP: {slide_info['tip']}"
            rtip.font.size = Pt(11)
            rtip.font.bold = True
            rtip.font.color.rgb = c_gold
            rtip.space_before = Pt(8)

        # Footer
        ft = s.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.7), Inches(0.35))
        fp = ft.text_frame.paragraphs[0]
        fp.text = f"TERRESTRIAL PLATFORM  |  {title_deck.upper()}  |  CONFIDENTIAL - FOR AUTHORISED PRESENTERS & SYSTEM INTEGRATORS"
        fp.font.size = Pt(9.5)
        fp.font.color.rgb = c_muted

        # Add Speaker Notes into PowerPoint's native notes slide
        notes_frame = s.notes_slide.notes_text_frame
        notes_text = f"--- SPEECH SCRIPT ---\n\n" + "\n\n".join(slide_info['speech_script'])
        notes_text += "\n\n--- TALKING POINTS ---\n" + "\n".join([f"- {p}" for p in slide_info['talking_points']])
        if 'tip' in slide_info:
            notes_text += f"\n\nPRESENTER TIP: {slide_info['tip']}"
        notes_frame.text = notes_text

    for out_path in output_paths:
        prs.save(out_path)
        print(f"Saved deck to: {out_path}")

# ==========================================================
# 1. EXECUTIVE PRESENTATION SPEECH DECK DATA
# ==========================================================
exec_slides = [
    {
        "step_tag": "OPENING HOOK | 0-60 SECONDS",
        "title": "The Industrial Dilemma: Why SCADA Is Failing Modern Plants",
        "speech_script": [
            "Good morning gentlemen and members of the board. Today, heavy industry faces a quiet crisis of productivity.",
            "Your power plants, refineries, and manufacturing facilities generate hundreds of thousands of sensor readings every single second. Yet over 80% of that data sits completely cold in historians.",
            "When a multi-million-dollar turbine trips or a chemical distillation column drifts off-spec, the damage is already done. Today, we are not introducing another monitoring dashboard. We are introducing the industrial nervous system: Terrestrial Platform."
        ],
        "talking_points": [
            "80%+ of plant telemetry sits unmonitored in historians",
            "Reactive alarms trigger only after physical damage occurs",
            "Terrestrial Platform transforms passive sensors into closed-loop autonomy",
            "Target Audience: CEOs, CFOs, COOs, and Plant Directors"
        ],
        "tip": "Speak with steady, grounded authority. Pause after '80% of data sits completely cold'."
    },
    {
        "step_tag": "CORE SHIFT | 2-4 MINUTES",
        "title": "From Reactive Alarms to Deterministic Closed-Loop Autonomy",
        "speech_script": [
            "For thirty years, industry has accepted static SCADA. If a bearing temperature exceeds 95 degrees, an alarm rings. That is not intelligence; that is an obituary for your machinery.",
            "Terrestrial Platform delivers true Level 4 Supervisory Autonomy. Instead of waiting for an alarm, our physics-informed models anticipate multi-variable drift up to 72 hours in advance.",
            "More importantly, we don't just alert your operators—our edge engine calculates optimal supervisory setpoints and safely closes the loop in sub-25 milliseconds."
        ],
        "talking_points": [
            "Traditional SCADA is merely a post-mortem indicator",
            "72-hour multi-variable predictive trajectory lookahead",
            "Sub-25 millisecond automated supervisory closed-loop actuation",
            "Operators transition from reactive fire-fighters to supervisory governors"
        ],
        "tip": "Contrast the word 'Obituary' with 'Autonomous Anticipation' to emphasize the mindset shift."
    },
    {
        "step_tag": "FINANCIAL IMPACT | 4-7 MINUTES",
        "title": "Quantifiable C-Suite ROI: EBITDA Expansion & Cost Compression",
        "speech_script": [
            "Let us talk about numbers that directly impact your quarterly financial statements.",
            "Across energy, utilities, and continuous process plants, Terrestrial Platform delivers an immediate 2% to 5% reduction in specific fuel and power consumption.",
            "We reduce unplanned downtime events by up to 70%, increase overall asset lifetime by 22%, and free up working capital. For a 500-megawatt generation unit or medium refinery, this translates to $3.2 Million to $8.5 Million in recurring annual EBITDA expansion."
        ],
        "talking_points": [
            "2% - 5% verified specific fuel and power consumption reduction",
            "Up to 70% reduction in catastrophic unplanned outages",
            "22% extension in high-capital rotating asset operational lifespan",
            "Verifiable project payback achieved within 4 to 6 months"
        ],
        "tip": "CFOs lean in when you mention EBITDA expansion. Show confidence on the 4-6 month payback."
    },
    {
        "step_tag": "TECHNICAL MOAT | 7-10 MINUTES",
        "title": "The Physics Moat: Why General LLMs Fail and PINO Wins",
        "speech_script": [
            "Many software vendors will try to pitch you generic artificial intelligence or generative LLMs for your industrial plant. Let us be unequivocal: Generative LLMs hallucinate.",
            "A hallucination on a consumer smartphone is an annoyance. A hallucination on a 1,000-degree boiler or high-voltage substation is an industrial catastrophe.",
            "Terrestrial Platform embeds first-principles physics laws—Navier-Stokes fluid dynamics, ASME steam tables, and ISO vibration thresholds—directly into our Neural Operators. If an actuation violates conservation of energy, the physics layer mathematically rejects it."
        ],
        "talking_points": [
            "Statistical LLMs hallucinate and lack physical ground truth",
            "Physics-Informed Neural Operators (PINO) enforce thermodynamic laws",
            "Strict adherence to ASME, ISO 10816, and IEC safety standards",
            "Deterministic, mathematically bounded actuation guardrails"
        ],
        "tip": "Use this slide to eliminate competition from generic tech giants who lack heavy industrial domain."
    },
    {
        "step_tag": "DEPLOYMENT | 10-12 MINUTES",
        "title": "Zero Rip-and-Replace: Live Onboarding in Under 48 Hours",
        "speech_script": [
            "Your Chief Operating Officer's biggest hesitation is always: 'Will this shut down my plant during installation?' The answer is a categorical no.",
            "Terrestrial Platform is completely non-intrusive. We deploy as a containerized edge runtime onto an existing industrial PC or VM. We interface non-intrusively with your existing Honeywell, Yokogawa, Siemens, or ABB DCS over OPC UA and Modbus.",
            "There is zero hardware rewiring, zero DCS logic modification, and zero plant downtime. We ingest and normalize your plant telemetry in under 48 hours."
        ],
        "talking_points": [
            "100% non-intrusive containerized software edge deployment",
            "Native interoperability with all legacy DCS and PLC vendors",
            "Zero plant shutdown required during installation and commissioning",
            "Historian telemetry normalized and operational in 48 hours"
        ],
        "tip": "Address the fear of downtime immediately. Emphasize 'Zero Rip-and-Replace'."
    },
    {
        "step_tag": "CALL TO ACTION | CLOSING 2 MINUTES",
        "title": "The 48-Hour Autonomous Pilot: Low Risk, High Verifiability",
        "speech_script": [
            "We do not ask for a multi-million-dollar leap of faith. We believe enterprise software should earn its trust on the plant floor.",
            "We propose a bounded 48-hour pilot on a single critical asset train—be it your primary boiler, feed pump, gas turbine, or chiller system.",
            "We will connect in two days, run in shadow mode for two weeks, and verify the quantifiable efficiency gains directly against your existing historian data. Let us start with one asset and let the numbers speak for themselves."
        ],
        "talking_points": [
            "Bounded pilot on a single critical asset train",
            "48-hour turnkey onboarding in shadow observation mode",
            "Direct comparison against historical baseline data",
            "Clear go/no-go milestone review at Day 14"
        ],
        "tip": "End with clear momentum. Hand over the one-page Pilot Agreement template."
    },
    {
        "step_tag": "BATTLECARD | EXECUTIVE Q&A",
        "title": "Anticipated Boardroom Objections & Strategic Rebuttals",
        "speech_script": [
            "Question: 'Why can't our current DCS vendor (Siemens/Honeywell) do this?' Rebuttal: DCS vendors sell proprietary hardware controllers. Their business model depends on vendor lock-in. Terrestrial Platform is hardware-agnostic and unifies all vendor silos with neural physics models.",
            "Question: 'What about cybersecurity in air-gapped critical infrastructure?' Rebuttal: Terrestrial Edge runs 100% on-premise without requiring outbound internet access. It conforms to IEC 62443 cyber architecture.",
            "Question: 'What if an autonomous setpoint causes an issue?' Rebuttal: Level 1 hardwired Safety Instrumented Systems (SIS) always take precedence, and operators maintain instantaneous bumpless manual override."
        ],
        "talking_points": [
            "DCS Vendor Lock-in Rebuttal: Agnostic across ABB, Siemens, Emerson, Yokogawa",
            "Air-Gap Security Rebuttal: Full on-premise execution with zero internet dependency",
            "Safety Interlock Rebuttal: SIS safety loop always retains supreme hardwired override"
        ],
        "tip": "Keep these rebuttals in mind during the Q&A segment after your presentation."
    }
]

# ==========================================================
# 2. AUTOMATION & ENGINEERING TECHNICAL SPEECH DECK DATA
# ==========================================================
eng_slides = [
    {
        "step_tag": "TECHNICAL HOOK | 0-90 SECONDS",
        "title": "The Automation Dilemma: SCADA Saturation & Alarm Flooding",
        "speech_script": [
            "Good morning automation leaders and engineering teams. Let's speak plainly about the operational reality inside our control rooms.",
            "During a plant upset, an operator doesn't receive two or three actionable notifications. They are hit with 1,500 alarms flooding across three monitors in under 90 seconds. Cryptic tag IDs, blinking red lights, and contradictory readings.",
            "Traditional SCADA is a 40-year-old paradigm designed to display raw voltage and register values. It lacks semantic plant context, thermodynamics, and causal relationship awareness. Terrestrial Platform was built by OT engineers to fix this fundamental architecture."
        ],
        "talking_points": [
            "Alarm floods of 1,000+ alarms/min overwhelm operator cognitive limits",
            "Cryptic tag dictionaries create silos between DCS, PLC, and vibration monitors",
            "SCADA provides static telemetry but zero first-principles causal inference",
            "Engineered specifically for OT Leads, DCS Programmers, and Plant Engineers"
        ],
        "tip": "Acknowledge the engineer's daily frustration. Establish yourself as a fellow engineer, not a marketer."
    },
    {
        "step_tag": "EDGE ARCHITECTURE | 2-5 MINUTES",
        "title": "Deterministic Edge Ingestion: OPC UA TSN, IEC 61850 & Modbus",
        "speech_script": [
            "Let's look at the telemetry ingestion layer. The Terra Edge Core runs as an ultra-lightweight Docker or Podman container on industrial PCs or hypervisor VMs.",
            "We ingest from your plant bus via native industrial protocol drivers: OPC UA with Time-Sensitive Networking (TSN), IEC 61850 MMS and GOOSE, Modbus TCP/RTU, DLMS/COSEM, and COMTRADE transient recorders.",
            "Ingestion occurs at deterministic sub-millisecond intervals. Every packet is microsecond-timestamped at arrival, parsed, normalized into our unified semantic asset model, and pushed to the local zero-copy shared memory bus."
        ],
        "talking_points": [
            "Zero-copy shared memory architecture for microsecond telemetry throughput",
            "Native support for OPC UA TSN, IEC 61850 GOOSE, Modbus, and COMTRADE",
            "Microsecond precision timestamping eliminates historian jitter",
            "Runs locally on existing industrial IPC or redundant edge VMs"
        ],
        "tip": "Highlight that this requires NO modification to existing PLC ladder logic or DCS function blocks."
    },
    {
        "step_tag": "CLOSED-LOOP AUTONOMY | 5-8 MINUTES",
        "title": "Deterministic Closed-Loop Supervisory Actuation & Fail-Safe Watchdogs",
        "speech_script": [
            "The question every automation engineer asks is: 'How does Terrestrial Platform safely write setpoints back into our DCS?'",
            "We operate strictly in Level 2 / Level 3 supervisory setpoint optimization mode. Our physics-informed engine calculates the optimal multi-variable trajectory—such as air-fuel ratio, boiler feed pump speed, or chiller staging.",
            "Before any setpoint is dispatched, it must pass a three-stage deterministic check: rate-of-change clamping, metallurgical stress limits, and safety boundary interlocks. If our heartbeat watchdog detects even a 50ms latency anomaly, it executes an instantaneous bumpless fallback to baseline manual setpoints."
        ],
        "talking_points": [
            "Supervisory setpoint writes via standard OPC UA DA/HA or Modbus registers",
            "Rate-of-change (RoC) and thermodynamic boundary clamping guardrails",
            "Hardware heartbeat watchdog with 50ms bumpless fail-safe fallback",
            "Level 1 Safety Instrumented Systems (SIS/ESD) retain ultimate hardwired authority"
        ],
        "tip": "Reassure them on safety: The ESD/SIS system always wins. We operate purely in supervisory setpoint mode."
    },
    {
        "step_tag": "PHYSICS MODELING | 8-11 MINUTES",
        "title": "Thermodynamic & Mechanical First-Principles Grounding (PINO)",
        "speech_script": [
            "How do our models achieve superior accuracy compared to standard regression? We ground our Neural Operators in first-principles industrial science.",
            "For thermal assets, we embed ASME IAPWS-IF97 steam tables directly into the computation graph. For rotating machinery, we embed ISO 10816 vibration severity standards and spectral FFT harmonic analysis.",
            "This means the model does not just look for statistical correlations; it understands mass conservation, enthalpy balance, and rotor dynamics. When ambient temperature swings by 20 degrees, the model adapts deterministically without retraining."
        ],
        "talking_points": [
            "IAPWS-IF97 ASME Steam Tables embedded for thermal power and chemical boilers",
            "ISO 10816 vibration envelopes & spectral peak tracking for rotating trains",
            "Mass, momentum, and energy conservation enforced via loss function penalties",
            "Eliminates statistical drift and ensures deterministic physics accuracy"
        ],
        "tip": "Engineers respect specific standards. Mentioning IAPWS-IF97 and ISO 10816 builds instant credibility."
    },
    {
        "step_tag": "ALARM INTELLIGENCE | 11-13 MINUTES",
        "title": "Causal Graph Alarm Suppression: Eliminating Nuisance Floods",
        "speech_script": [
            "Let's solve the alarm flood problem. Terrestrial Platform's TerraAlarm module imports your piping and instrumentation diagrams (P&ID) and electrical single-line diagrams to build a Directed Causal Graph.",
            "When an upstream pump cavitation occurs, our engine detects the propagation wavefront across discharge pressure, suction flow, and motor current.",
            "In under 50 milliseconds, it suppresses all 45 secondary symptom alarms and highlights only the primary root cause: 'Pump A Suction Strainer Clogged'. Operators know immediately where to send technicians."
        ],
        "talking_points": [
            "Directed causal graph mapped directly to plant P&ID topology",
            "Sub-50 millisecond propagation wavefront detection",
            "Suppresses 95%+ of nuisance cascading alarms during plant trips",
            "Conforms strictly to ISA 18.2 and EEMUA 191 alarm management standards"
        ],
        "tip": "Reference ISA 18.2 compliance. Plant automation engineers are judged heavily on alarm metrics."
    },
    {
        "step_tag": "CYBERSECURITY | 13-15 MINUTES",
        "title": "OT/IT Air-Gap Architecture & IEC 62443 Compliance",
        "speech_script": [
            "We adhere rigorously to the Purdue Enterprise Reference Model (ISA-95).",
            "The Terra Edge node resides securely in Level 2.5 / Level 3 DMZ. It requires zero outbound internet connectivity to function and never touches Level 1 safety controllers directly.",
            "All internal communication utilizes TLS 1.3 encryption, mTLS certificate authentication, and strict Role-Based Access Control (RBAC). The platform includes complete 21 CFR Part 11 compliant audit logging for every supervisory actuation."
        ],
        "talking_points": [
            "Strict adherence to Purdue Model (ISA-95 Level 2.5 / 3 boundary)",
            "100% operational in air-gapped industrial enclaves with zero external WAN",
            "Mutual TLS 1.3 encryption and granular Role-Based Access Control (RBAC)",
            "Conforms to IEC 62443 and 21 CFR Part 11 immutable audit logging"
        ],
        "tip": "Cybersecurity compliance is non-negotiable for OT engineers. Reassure them that air-gap is respected."
    },
    {
        "step_tag": "DEPLOYMENT | 15-18 MINUTES",
        "title": "The Turnkey 48-Hour Engineering Deployment Roadmap",
        "speech_script": [
            "Here is the exact rollout playbook your engineering team will experience.",
            "Day 1: Container spin-up on your designated industrial workstation or virtual machine. Automated import of DCS tag list via CSV or SCL files.",
            "Day 2: Protocol handshake verification over OPC UA / Modbus. Ingestion validation against existing DCS trend screens.",
            "Days 3 to 14: Shadow calibration mode. The models run alongside your operators without writing setpoints, verifying forecast accuracy. Day 15: Enable advisory mode, followed by closed-loop supervisory execution."
        ],
        "talking_points": [
            "Day 1: Container launch & automated DCS tag dictionary mapping",
            "Day 2: Protocol handshake and multi-variable signal validation",
            "Day 3-14: Passive shadow calibration against historian trends",
            "Day 15+: Staged transition to advisory warnings and closed-loop actuation"
        ],
        "tip": "A clear chronological timeline removes the fear of open-ended implementation headaches."
    },
    {
        "step_tag": "TECHNICAL FAQ | Q&A SEGMENT",
        "title": "Automation Engineer Technical Battlecard & FAQ",
        "speech_script": [
            "Question: 'Will this increase scan cycle time on our PLC/DCS controllers?' Rebuttal: No. All AI inference and physics calculation happens on the edge PC runtime. We communicate with the PLC purely through standard asynchronous register polling, consuming less than 1% PLC CPU.",
            "Question: 'Can it connect to our historical database (OSIsoft PI / Wonderware / IP.21)?' Rebuttal: Yes. Terra Edge includes native connectors for PI Web API, OPC HDA, TimescaleDB, and standard SQL historians.",
            "Question: 'What happens if the edge PC crashes or loses power?' Rebuttal: The DCS watchdog immediately detects loss of heartbeat and maintains the last known safe manual setpoint with zero bump."
        ],
        "talking_points": [
            "PLC CPU Impact: Under 1% overhead via standard asynchronous polling",
            "Historian Compatibility: Native OSIsoft PI, Aveva Wonderware, Aspentech IP.21",
            "Power/Hardware Loss: Instantaneous fail-safe watchdog transfer with zero bump"
        ],
        "tip": "Answer technical questions with exact protocol names and scan-time numbers."
    }
]

# Generate Executive Deck
create_deck(
    title_deck="Executive Pitch & Speech Script Guide",
    deck_tag="EXECUTIVE PERSONNEL",
    slides_data=exec_slides,
    output_paths=[
        os.path.join(website_dir, "Executive_Speech_Deck.pptx"),
        os.path.join(meta_dir, "Executive_Speech_Deck.pptx")
    ]
)

# Generate Engineering Deck
create_deck(
    title_deck="Automation & Engineering Technical Speech Guide",
    deck_tag="AUTOMATION & ENGINEERING",
    slides_data=eng_slides,
    output_paths=[
        os.path.join(website_dir, "Engineering_Speech_Deck.pptx"),
        os.path.join(meta_dir, "Engineering_Speech_Deck.pptx")
    ]
)

print("SUCCESS: Both Executive and Engineering speech decks created successfully!")


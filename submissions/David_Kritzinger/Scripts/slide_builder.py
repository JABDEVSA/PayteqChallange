from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

OUTPUT = Path("build-output")
OUTPUT.mkdir(exist_ok=True)

ASSETS = Path("presentation-assets")
ASSETS.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT / "Wedding-Planner-Proposal.pptx"

# Optional screenshots. Missing files are replaced with labelled panels.
DASHBOARD_IMAGE = ASSETS / "dashboard.png"
BUDGET_IMAGE = ASSETS / "budget.png"
SEATING_IMAGE = ASSETS / "seating.png"

# Replace these with the exact wording from validation.md.
TEST_ENVIRONMENT = "INSERT EXACT TESTED ENVIRONMENT FROM validation.md"
VALIDATION_SUMMARY = (
    "INSERT VERIFIED TEST OUTCOMES AND ANY UNRESOLVED ISSUES."
)

COLOURS = {
    "ivory": "FAF8F2",
    "forest": "344A3B",
    "sage": "81917A",
    "pale": "EDF2E9",
    "blush": "E9D1C8",
    "amber": "FFF3D6",
    "muted": "637064",
    "white": "FFFFFF",
    "line": "D9DED5",
}

FONT = "Aptos"
TITLE_FONT = "Georgia"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


# ------------------------------------------------------------
# Layout helpers
# ------------------------------------------------------------

def rgb(value):
    return RGBColor.from_string(COLOURS.get(value, value))


def rectangle(slide, x, y, w, h, fill, line=None, rounded=False):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line:
        shape.line.color.rgb = rgb(line)
    else:
        shape.line.fill.background()
    if rounded:
        shape.adjustments[0] = 0.08
    return shape


def textbox(slide, x, y, w, h, text, size=18, colour="forest",
            bold=False, font=FONT):
    box = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)

    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = font
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = rgb(colour)
        p.space_after = Pt(9)
    return box


def base_slide(title, subtitle="", status="PROPOSED"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb("ivory")

    rectangle(slide, 0, 0, 13.333, 0.11, "sage")
    textbox(slide, 0.55, 0.35, 10.8, 0.65,
            title, 29, bold=True, font=TITLE_FONT)

    if subtitle:
        textbox(slide, 0.58, 1.06, 12.0, 0.62,
                subtitle, 15, colour="muted")

    rectangle(slide, 11.5, 0.42, 1.25, 0.34, "pale", rounded=True)
    textbox(slide, 11.56, 0.46, 1.15, 0.23,
            status, 9, bold=True)

    rectangle(slide, 0.55, 7.03, 12.2, 0.012, "line")
    textbox(
        slide, 0.58, 7.12, 11.5, 0.2,
        "Assessment adaptation of Johan Botha’s reference pack • "
        "Commercial proposals require authorisation",
        8, colour="muted"
    )
    textbox(slide, 12.15, 7.09, 0.5, 0.25,
            f"{len(prs.slides):02}", 10, colour="muted")
    return slide


def card(slide, x, y, w, h, title, body, fill="white", size=18):
    rectangle(slide, x, y, w, h, fill, "line", rounded=True)
    textbox(slide, x + 0.2, y + 0.17, w - 0.4, 0.45,
            title, 20, bold=True)
    textbox(slide, x + 0.2, y + 0.8, w - 0.4, h - 0.95,
            body, size)


def banner(slide, text, y=6.27, fill="pale", size=17):
    rectangle(slide, 0.58, y, 12.16, 0.54, fill, rounded=True)
    textbox(slide, 0.76, y + 0.1, 11.8, 0.4, text, size, bold=True)


def image_panel(slide, path, x, y, w, h, placeholder):
    rectangle(slide, x, y, w, h, "white", "line", rounded=True)

    if path.exists():
        # Fit the full image within the panel without distortion.
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y))
        scale = min(Inches(w) / pic.width, Inches(h) / pic.height)
        pic.width = int(pic.width * scale)
        pic.height = int(pic.height * scale)
        pic.left = Inches(x) + (Inches(w) - pic.width) // 2
        pic.top = Inches(y) + (Inches(h) - pic.height) // 2
    else:
        textbox(
            slide, x + 0.25, y + 0.5, w - 0.5, h - 0.7,
            placeholder, 19, colour="muted"
        )


def table(slide, headers, rows, x, y, w, h, widths=None, size=15):
    shape = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    t = shape.table

    if widths:
        for col, fraction in zip(t.columns, widths):
            col.width = Inches(w * fraction)

    for r, values in enumerate([headers] + rows):
        for c, value in enumerate(values):
            cell = t.cell(r, c)
            cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(
                "forest" if r == 0 else ("white" if r % 2 else "pale")
            )
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.08)
            cell.margin_bottom = Inches(0.06)

            for p in cell.text_frame.paragraphs:
                p.font.name = FONT
                p.font.size = Pt(size)
                p.font.bold = r == 0
                p.font.color.rgb = rgb("white" if r == 0 else "forest")
    return shape


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ------------------------------------------------------------
# 01 — Recommendation
# ------------------------------------------------------------

s = base_slide(
    "A working planner. A practical growth path.",
    "Wedding planning in Excel → clearer offers → measured business expansion",
    "BUILT + PROPOSED"
)

textbox(
    s, 0.62, 1.95, 5.0, 2.0,
    "Strengthen the downloadable product first.\n"
    "Add online services only when customers justify them.",
    29, bold=True, font=TITLE_FONT
)

textbox(
    s, 0.65, 4.25, 4.8, 1.55,
    "• 21 planning areas\n"
    "• Clean master + two client demonstrations\n"
    "• Personal and professional offers\n"
    "• No mandatory subscription",
    20
)

image_panel(
    s, DASHBOARD_IMAGE, 5.95, 1.85, 6.75, 4.2,
    "Add dashboard.png to presentation-assets/\n\n"
    "Use a screenshot from the recalculated Excel workbook."
)

banner(s, "Recommendation: validate the packaged Excel offer before funding a full application.")

notes(s,
    "Open with the decision, not the implementation language. "
    "The workbook is implemented; prices, operations and future software "
    "are proposals for the project owner."
)


# ------------------------------------------------------------
# 02 — Buyers
# ------------------------------------------------------------

s = base_slide(
    "Two buyers, two different jobs",
    "The same core workbook creates different value—and needs different permissions."
)

table(s,
    ["", "Couple / personal buyer", "Professional / agency"],
    [
        ["Problem", "Organise one wedding", "Repeat a reliable process across clients"],
        ["Value", "Clarity, affordability and control", "Time saved and consistent delivery"],
        ["Use", "One wedding and its planning team", "Separate workbook for each client"],
        ["Support", "Simple setup and troubleshooting", "Reuse, handover and version guidance"],
        ["Buying pattern", "Usually once-off", "Repeat use and licence expansion"],
    ],
    0.6, 1.9, 12.1, 3.65,
    widths=[0.18, 0.39, 0.43], size=17
)

banner(s, "Positioning: more structured than a blank spreadsheet; simpler than a compulsory online system.")

notes(s,
    "Do not claim Excel is better for everyone. The differentiation is an "
    "offline, structured, reusable workflow. Professional usage is not "
    "automatically included in a personal purchase."
)


# ------------------------------------------------------------
# 03 — Coverage
# ------------------------------------------------------------

s = base_slide(
    "The product covers all 21 planning areas",
    "Original spreadsheet functionality—not screenshots pasted into cells.",
    "IMPLEMENTED"
)

groups = [
    ("Overview & setup", "Dashboard • Setup"),
    ("Style & announcements", "Save the Date • Theme Board"),
    ("Scheduling & logistics",
     "Calendar • Planning Timeline\nWedding Itinerary • Packing List"),
    ("Suppliers & money",
     "Vendor Choices • Venue Options\nWedding Budget • Contact Directory"),
    ("People & hospitality",
     "Guest List • Seating Plan\nWedding Party • Food and Drinks"),
    ("Memories & follow-through",
     "Photoshoot • Photo Gallery • Music\nGifts and Thank You • Honeymoon Budget"),
]

for i, (heading, body) in enumerate(groups):
    col = i % 3
    row = i // 3
    card(s, 0.6 + col * 4.1, 1.85 + row * 1.92,
         3.95, 1.75, heading, body, size=16)

banner(
    s,
    "Wedding-Planner.xlsx = clean master  |  Client-A.xlsx + Client-B.xlsx = demonstrations",
    size=16
)

notes(s,
    "The full screenshot-to-sheet map belongs in README.md. "
    "Start Here and the hidden Lists sheet support the 21 planning areas."
)


# ------------------------------------------------------------
# 04 — Functionality
# ------------------------------------------------------------

s = base_slide(
    "Inputs become useful planning decisions",
    "Connected calculations support money, guests and scheduling.",
    "IMPLEMENTED"
)

card(s, 0.6, 1.85, 3.95, 2.6,
     "Money",
     "Budget entries\n↓\n"
     "Planned, paid and outstanding\n↓\n"
     "Dashboard + payment calendar",
     fill="pale", size=18)

card(s, 4.7, 1.85, 3.95, 2.6,
     "Guests",
     "RSVPs + table assignments\n↓\n"
     "Confirmed attendance\n↓\n"
     "Seating views + spaces left",
     fill="pale", size=18)

card(s, 8.8, 1.85, 3.95, 2.6,
     "Planning",
     "Wedding date + task rules\n↓\n"
     "Calculated due dates\n↓\n"
     "Calendar + ordered open tasks",
     fill="pale", size=18)

textbox(
    s, 0.75, 4.8, 11.9, 1.0,
    "Amber = editable inputs     •     Pale green = calculations\n"
    "Native Excel Tables grow with records; Guest List is the source of truth for seating.",
    20
)

banner(s, "No account, external database, macros or Python required to use the workbook.")

notes(s,
    "Fixed task dates override relative dates. Menu quantities are entered "
    "manually, not automatically generated from RSVPs. Python generates the "
    "files but is not needed by the purchaser."
)


# ------------------------------------------------------------
# 05 — Validation
# ------------------------------------------------------------

s = base_slide(
    "Validation: show the workbook working",
    "Populate this summary from the completed validation.md—not from code inspection.",
    "EVIDENCE REQUIRED"
)

table(s,
    ["Area", "Evidence to reference"],
    [
        ["Budget and payments", "Changed costs/payments → correct totals and remaining budget"],
        ["Guests and seating", "Changed RSVP/table → consistent counts and seating views"],
        ["Tasks and menu", "Date dependencies; fixed-date override; quantity × unit cost"],
        ["Growth and blanks", "Added records and cleared optional inputs"],
        ["Persistence and copies", "Saved/reopened; one client changed without affecting the other"],
        ["Clean start and print", "Reusable master; readable budget and itinerary output"],
    ],
    0.6, 1.85, 12.1, 3.5,
    widths=[0.25, 0.75], size=16
)

textbox(s, 0.65, 5.55, 12.0, 0.45,
        f"Environment: {TEST_ENVIRONMENT}", 15, bold=True)

banner(s, VALIDATION_SUMMARY, fill="amber", size=15)

notes(s,
    "Replace all placeholders before submission. Add actual outcomes, any "
    "failed tests and references to screenshots in validation.md. "
    "Use only the environment personally tested."
)


# ------------------------------------------------------------
# 06 — Journey
# ------------------------------------------------------------

s = base_slide(
    "Design the complete customer experience",
    "A good digital product includes the journey around the file."
)

table(s,
    ["Stage", "Proposed experience"],
    [
        ["Discover", "Accurate previews, short demonstration and compatibility requirements"],
        ["Purchase", "Clear personal/professional usage terms before payment"],
        ["Download", "Workbook, quick-start guide, licence and version information"],
        ["First use", "Save As → enter Setup details → start with amber cells"],
        ["Planning", "Record costs, guests and tasks; review the dashboard"],
        ["Wedding day", "Print or export itinerary and relevant lists"],
        ["Afterward", "Complete thank-you tracking; archive securely or start a fresh client copy"],
    ],
    0.6, 1.85, 12.1, 4.1,
    widths=[0.18, 0.82], size=16
)

banner(s, "Professionals start each client from the clean master—not a previous client’s workbook.")

notes(s,
    "A suitable Excel installation is required. Standalone does not mean "
    "Excel itself is free. Existing customer files do not automatically "
    "migrate when a new template version is released."
)


# ------------------------------------------------------------
# 07 — Offers
# ------------------------------------------------------------

s = base_slide(
    "Sell clear usage rights—not vague access",
    "Illustrative ZAR prices and terms to test; not verified market pricing."
)

table(s,
    ["Offer", "Test price", "Proposed permission"],
    [
        ["Personal", "R299 once-off", "One wedding; share with that wedding’s planning team"],
        ["Professional", "R1,499 once-off", "One named planner; up to 10 client weddings"],
        ["Agency", "R3,999 once-off", "Up to five staff; up to 30 client weddings"],
        ["Customisation", "Quoted", "Scoped branding, onboarding or process changes"],
    ],
    0.6, 1.85, 12.1, 2.65,
    widths=[0.2, 0.19, 0.61], size=17
)

textbox(
    s, 0.75, 4.75, 11.8, 1.1,
    "Proposed terms: downloaded files do not expire; retain completed client files.\n"
    "No resale or redistribution of the blank master. Additional capacity requires an upgrade.\n"
    "Test a 12-month update entitlement and 30-day onboarding support; scope custom work separately.",
    17
)

banner(s, "These are proposed customer terms—not commercial rights granted by this assessment.")

notes(s,
    "No licence server or intrusive DRM is proposed. Manual entitlement "
    "tracking is adequate initially. Update entitlement does not imply "
    "automatic changes to existing populated workbooks. Confirm legal "
    "wording and support feasibility before launch."
)


# ------------------------------------------------------------
# 08 — Economics
# ------------------------------------------------------------

s = base_slide(
    "Revenue is not profit",
    "Illustrative monthly scenario—not a forecast or claim of existing sales."
)

card(
    s, 0.6, 1.85, 4.1, 3.8,
    "Assumed monthly sales",
    "40 personal × R299\n= R11,960\n\n"
    "4 professional × R1,499\n= R5,996\n\n"
    "Gross revenue: R17,956",
    fill="pale", size=22
)

table(s,
    ["Assumed cost", "Amount"],
    [
        ["Selling/payment fees: 12%", "R2,155"],
        ["Promotion", "R3,000"],
        ["Support: 8h × R200", "R1,600"],
        ["Maintenance: 6h × R250", "R1,500"],
        ["Tools and administration", "R500"],
        ["Contribution before tax / other costs", "R9,201"],
    ],
    4.95, 1.85, 7.75, 3.8,
    widths=[0.76, 0.24], size=17
)

banner(s, "Validate margin after acquisition and support—not sales volume alone.")

notes(s,
    "Figures are rounded to the nearest rand. Actual contribution is "
    "R9,201.28 with exactly 12% fees. Verify platform fees, taxes and "
    "refund exposure. Initial development, legal/setup costs and other "
    "overheads are excluded. Professional support may exceed this assumption."
)


# ------------------------------------------------------------
# 09 — Acquisition
# ------------------------------------------------------------

s = base_slide(
    "Promote the outcome; measure the economics",
    "Proposed activities only—no live listings, customer contact or advertising for this assessment."
)

card(
    s, 0.6, 1.85, 4.1, 3.95,
    "Prioritise three activities",
    "1. Better previews and a short workbook walkthrough.\n\n"
    "2. Helpful budget, RSVP and timeline content.\n\n"
    "3. Authorised planner interviews and a small professional pilot.",
    fill="pale", size=19
)

table(s,
    ["Measure", "Decision"],
    [
        ["Purchase conversion", "Improve previews, positioning or offer"],
        ["Acquisition cost", "Continue or stop a channel"],
        ["Support minutes/order", "Improve onboarding or pricing"],
        ["Refund reasons", "Fix expectations and usability"],
        ["Licence expansion", "Assess professional repeat value"],
        ["Contribution/order", "Assess sustainability"],
    ],
    4.95, 1.85, 7.75, 3.95,
    widths=[0.43, 0.57], size=16
)

banner(s, "Creative concept: “One wedding. One clear plan. No mandatory subscription.”")

notes(s,
    "The creative message is a proposed concept, not a published ad. "
    "Avoid assuming a couple is a recurring subscriber. Professional "
    "capacity upgrades and referrals may matter more than consumer retention."
)


# ------------------------------------------------------------
# 10 — Operations
# ------------------------------------------------------------

s = base_slide(
    "A control system is needed—but not a custom platform first",
    "Buy commodity capabilities; document the process; automate proven bottlenecks."
)

table(s,
    ["Control", "Initial approach", "Why"],
    [
        ["Orders / delivery", "Marketplace or delivery service", "Avoid rebuilding checkout and fulfilment"],
        ["Customers / licences", "Restricted structured register", "Track edition and usage entitlement"],
        ["Versions / updates", "Release archive + changelog", "Know which product each buyer received"],
        ["Support / communication", "Shared inbox or help desk", "Track issues and permitted update messages"],
        ["Payments / billing", "Provider records + accounting", "Reconcile sales, fees, refunds and settlements"],
        ["Reporting / admin", "Monthly dashboard + checklist", "Monitor economics and reduce release errors"],
    ],
    0.6, 1.85, 12.1, 3.9,
    widths=[0.2, 0.34, 0.46], size=16
)

banner(s, "Build a small entitlement/support portal only when manual work becomes costly or unreliable.")

notes(s,
    "Use existing tools first. Assign an administrator for releases and "
    "permissions; restrict support access to the information needed. "
    "Do not collect wedding guest lists just to sell template licences."
)


# ------------------------------------------------------------
# 11 — Architecture
# ------------------------------------------------------------

s = base_slide(
    "Keep commerce records separate from wedding data",
    "Proposed architecture; the current Excel files do not upload or synchronise information."
)

# Flow blocks and simple directional labels.
card(s, 0.6, 1.9, 2.45, 1.45,
     "Buyer", "Purchases an edition", fill="pale", size=16)
textbox(s, 3.13, 2.4, 0.35, 0.4, "→", 26)

card(s, 3.5, 1.9, 2.65, 1.45,
     "Checkout", "Order + payment status", fill="pale", size=16)
textbox(s, 6.23, 2.4, 0.35, 0.4, "→", 26)

card(s, 6.6, 1.9, 2.65, 1.45,
     "Licence register", "Edition + entitlement", fill="pale", size=16)
textbox(s, 9.33, 2.4, 0.35, 0.4, "→", 26)

card(s, 9.7, 1.9, 3.0, 1.45,
     "Versioned delivery", "Local Excel workbook", fill="pale", size=16)

card(
    s, 0.6, 3.7, 5.9, 2.05,
    "Business-controlled records",
    "Customer / organisation • Order • Licence\n"
    "Product version • Support ticket • Payment reference\n"
    "Payment provider handles payment credentials.",
    size=17
)

card(
    s, 6.75, 3.7, 5.95, 2.05,
    "Customer-controlled wedding data",
    "Couple / planner stores workbook locally.\n"
    "Administrator manages releases; support gets limited access.\n"
    "MFA, backups, release checks and manual delivery fallback.",
    size=17
)

banner(s, "Use native integrations first; confirm API availability, permissions and cost before committing.")

notes(s,
    "Buyer journey: purchase, receive correct version, request support. "
    "Admin journey: release, fulfil, resolve, reconcile. Accounting consumes "
    "payment and refund references. The register should not hold card data. "
    "Buyers share files themselves; local copies cannot enforce fine-grained "
    "permissions. Operating tool cost begins with the R500 monthly assumption "
    "on slide 8 and must be validated."
)


# ------------------------------------------------------------
# 12 — Platforms
# ------------------------------------------------------------

s = base_slide(
    "Choose platforms by the problem they solve",
    "Excel is the starting product—not a temporary placeholder."
)

table(s,
    ["Option", "Decision", "Reason / trade-off"],
    [
        ["Excel download", "Keep as core", "Offline reuse; deliver updates as new versions with migration guidance"],
        ["PC desktop app", "Defer", "Needs value beyond Excel to justify installation, signing and support"],
        ["Web application", "Validate later", "Potential value in shared tasks and multi-client collaboration"],
        ["Native mobile app", "Defer", "Prove a specific on-the-go workflow; test responsive web first"],
        ["Professional workflow", "Start in Excel", "Client copies and reusable processes before multi-tenant software"],
    ],
    0.6, 1.85, 12.1, 3.65,
    widths=[0.21, 0.18, 0.61], size=17
)

textbox(
    s, 0.75, 5.65, 11.8, 0.45,
    "If a future hosted service ends: downloaded Excel files remain usable; define exports and hosted-data retention.",
    17, bold=True
)

notes(s,
    "Any web product adds authentication, permissions, backups, monitoring, "
    "privacy work and recurring operating cost. Excel lacks real-time "
    "synchronisation and granular access control. These are trade-offs, "
    "not reasons to build all platforms immediately."
)


# ------------------------------------------------------------
# 13 — Roadmap
# ------------------------------------------------------------

s = base_slide(
    "A gated 30 / 60 / 90-day plan",
    "Each phase earns the next investment."
)

card(
    s, 0.6, 1.85, 3.95, 3.85,
    "Days 1–30 | Package",
    "Buyer interviews; onboarding; licence review; offer design.\n\n"
    "Effort: product owner + limited legal/design input.\n\n"
    "Gate: credible problem and willingness to pay.",
    fill="pale", size=18
)

card(
    s, 4.7, 1.85, 3.95, 3.85,
    "Days 31–60 | Pilot",
    "Authorised pilot; price/offer tests; support and cost tracking.\n\n"
    "Effort: product owner + light technical support.\n\n"
    "Gate: usable product with manageable support burden.",
    fill="pale", size=18
)

card(
    s, 8.8, 1.85, 3.95, 3.85,
    "Days 61–90 | Decide",
    "Improve the proven offer; prototype one repeated workflow pain point.\n\n"
    "Effort: conditional development.\n\n"
    "Gate: paid demand for a specific online benefit.",
    fill="pale", size=18
)

banner(s, "Dependencies: owner permission, verified rules/fees, research consent and support capacity.")

notes(s,
    "Agree numerical success thresholds before the pilot. Do not invent "
    "benchmarks after seeing outcomes. This is a phased effort estimate, "
    "not a fixed-price implementation commitment."
)


# ------------------------------------------------------------
# 14 — Risks and recommendation
# ------------------------------------------------------------

s = base_slide(
    "Be explicit about limits—and the next decision",
    "The strongest recommendation is disciplined expansion, not maximum feature count."
)

card(
    s, 0.6, 1.85, 5.9, 3.95,
    "Current product boundaries",
    "• Suitable modern Excel required; only documented testing is claimed.\n"
    "• Formula cells are unlocked; paste can bypass validation.\n"
    "• Validation through row 5000; 20 tables; four gallery panels.\n"
    "• Cumulative payments—not an instalment ledger.\n"
    "• No FX conversion, live sync or automatic theme-wide restyling.\n"
    "• Images and print-area review remain manual.",
    size=17
)

card(
    s, 6.75, 1.85, 5.95, 3.95,
    "Business risks and responses",
    "Support burden → clearer onboarding and compatibility.\n\n"
    "Licence misuse → clear terms and upgrade paths.\n\n"
    "Privacy → minimal collection and careful file sharing.\n\n"
    "Premature spend → evidence gates before development.\n\n"
    "Ownership → authorisation before commercial use.",
    size=17
)

banner(s, "Next decision: approve an authorised Excel-offer pilot—not a full application build.")

notes(s,
    "Mention any additional limitations identified in validation.md. "
    "Do not hide failed tests. This assessment adaptation cannot be "
    "independently commercialised under the supplied assessment licence."
)


# ------------------------------------------------------------
# 15 — Sources and disclosure
# ------------------------------------------------------------

s = base_slide(
    "Evidence, assumptions and disclosure",
    "Complete source details and access dates before submission.",
    "REVIEW REQUIRED"
)

card(
    s, 0.6, 1.85, 5.9, 3.85,
    "Sources and evidence",
    "Supplied: assessment brief, reference pack and licence.\n\n"
    "Submission: workbook files, coverage map and validation.md.\n\n"
    "External: add only sources actually consulted, with URL and access date.\n\n"
    "Pricing, sales volume and cost figures are illustrative assumptions.",
    size=18
)

card(
    s, 6.75, 1.85, 5.95, 3.85,
    "AI and personal verification",
    "Disclose ChatGPT assistance with presentation structure and drafting.\n\n"
    "Add any other AI tools actually used for code, research or visuals.\n\n"
    "Describe your own changes and verification in ai-usage.md.\n\n"
    "Do not claim research or testing you did not perform.",
    size=18
)

banner(s, "Supporting files: README.md  •  validation.md  •  ai-usage.md  •  workbook/")

notes(s,
    "Suggested official research categories are marketplace fees and digital "
    "product rules, Excel function requirements and relevant privacy guidance. "
    "These are research tasks, not sources already verified by this deck."
)


# ------------------------------------------------------------
# Save
# ------------------------------------------------------------

prs.core_properties.title = "Wedding Planner — Product and Business Proposal"
prs.core_properties.subject = "Payteq candidate assessment"
prs.core_properties.author = "Candidate — replace with your name"
prs.core_properties.keywords = (
    "wedding planner, Excel, digital product, business proposal, assessment"
)
prs.core_properties.comments = (
    "Assessment adaptation of Johan Botha's supplied reference pack. "
    "Commercial proposals are illustrative and subject to authorisation."
)

prs.save(OUTPUT_FILE)

print(f"Created: {OUTPUT_FILE.resolve()}")
print()
print("Before submitting:")
print("1. Add actual workbook screenshots and rerun.")
print("2. Replace validation and environment placeholders.")
print("3. Complete sources, access dates and AI disclosure.")
print("4. Review every slide in PowerPoint for readability.")
print("5. Export the reviewed presentation as PDF.")
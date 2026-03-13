from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Inches(8.5)
section.page_height = Inches(11)
section.left_margin   = Inches(1.1)
section.right_margin  = Inches(1.1)
section.top_margin    = Inches(1.0)
section.bottom_margin = Inches(1.0)

# ── Colors ────────────────────────────────────────────────────────
NAVY   = RGBColor(0x1B, 0x3A, 0x6B)
GOLD   = RGBColor(0xC8, 0x96, 0x3E)
GREEN  = RGBColor(0x2D, 0x6A, 0x4F)
GRAY   = RGBColor(0x6B, 0x72, 0x80)
BLACK  = RGBColor(0x1A, 0x1A, 0x2E)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

# ── Helper: set paragraph shading ─────────────────────────────────
def shade_paragraph(para, fill_hex):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    pPr.append(shd)

def add_run(para, text, bold=False, color=None, size=None, italic=False):
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    if color: run.font.color.rgb = color
    if size:  run.font.size = Pt(size)
    return run

def h1(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = NAVY
    return p

def h2(text, color=NAVY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = color
    return p

def h3(num, label, badge):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(3)
    add_run(p, f"Step {num}  ", bold=True, color=NAVY, size=11.5)
    add_run(p, f"[{badge}]  ", bold=True, color=GOLD, size=9)
    add_run(p, label, bold=True, color=NAVY, size=11.5)
    return p

def body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.left_indent  = Inches(0.25)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(6)
    for run in p.runs:
        run.font.size = Pt(10.5)
        run.font.color.rgb = BLACK
    return p

def link_line(label, url):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_after = Pt(4)
    add_run(p, f"→ {label}: ", bold=True, color=GOLD, size=10)
    add_run(p, url, color=NAVY, size=10, italic=True)
    return p

def divider():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run("─" * 72)
    run.font.color.rgb = RGBColor(0xD1, 0xD5, 0xDB)
    run.font.size = Pt(8)
    return p

def section_banner(text, fill_hex, text_color=WHITE):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(10)
    shade_paragraph(p, fill_hex)
    run = p.add_run(f"  {text}  ")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = text_color
    return p

def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent   = Inches(0.4)
    p.paragraph_format.space_before  = Pt(1)
    p.paragraph_format.space_after   = Pt(3)
    if bold_prefix:
        add_run(p, bold_prefix + " ", bold=True, color=NAVY, size=10.5)
        add_run(p, text, color=BLACK, size=10.5)
    else:
        run = p.add_run(text)
        run.font.size = Pt(10.5)
        run.font.color.rgb = BLACK
    return p

# ══════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
shade_paragraph(p, "1B3A6B")
run = p.add_run("\n  Getting Started with Claude Code\n  A Complete Guide: Setup · Build · Deploy\n ")
run.bold = True
run.font.size = Pt(18)
run.font.color.rgb = WHITE

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(4)
add_run(p2, "Created with Claude Code  ·  anthropic.com", color=GRAY, size=9)

divider()

# ══════════════════════════════════════════════════════════════════
# PART 1: SETUP
# ══════════════════════════════════════════════════════════════════
section_banner("PART 1 — Set Up Claude Code", "1B3A6B")

intro = doc.add_paragraph(
    "This guide walks you through everything you need to get started with Claude Code — "
    "from creating your account to building and deploying a real website. "
    "No command line or prior coding experience required."
)
intro.paragraph_format.space_after = Pt(10)
for run in intro.runs:
    run.font.size = Pt(10.5)
    run.font.color.rgb = GRAY

# Step 1
h3(1, "Subscribe to Claude Pro", "Web")
body("Open your browser and go to claude.ai. Create a free account if you don't have one, "
     "then upgrade to the Pro plan for $20/month. This unlocks Claude's most capable models "
     "and gives you access to Claude Code.")
link_line("claude.ai", "https://claude.ai")

divider()

# Step 2
h3(2, "Install Claude Desktop", "Desktop")
body("Download the Claude Desktop app from claude.ai/download. Install it like any other app "
     "on your Mac or Windows computer. Sign in with your Anthropic account. Claude Desktop "
     "gives you a full AI assistant on your computer — great for chatting, brainstorming, "
     "and drafting content without opening a browser.")
link_line("Download Claude Desktop", "https://claude.ai/download")

divider()

# Step 3
h3(3, "Install Your Code Editor (IDE)", "IDE")
body("Download and install one of these AI-powered code editors — both are free and beginner-friendly:")
bullet("Antigravity (by Google) — download from the Google website and install like any app.")
bullet("Cursor — download from cursor.com and install like any app.")
body("Both editors have a visual, point-and-click interface. No command line needed.")
link_line("Download Cursor", "https://cursor.com")

divider()

# Step 4
h3(4, "Find the Claude Code Extension", "IDE")
body("Open your IDE (Antigravity or Cursor). In the left sidebar, click the Extensions icon "
     "(it looks like four small squares). In the search box, type Claude Code. "
     "Click the result and press Install. The extension connects your editor directly to Claude.")

divider()

# Step 5
h3(5, "Log In & Activate Claude Code", "IDE")
body("After installation, a Claude Code panel will appear in your IDE sidebar. Click it and "
     "sign in with your Anthropic account (the same login you use on claude.ai). "
     "You're now fully set up — Claude Code is ready to help you write, edit, and build.")

# ══════════════════════════════════════════════════════════════════
# PART 2: BUILD A WEBSITE
# ══════════════════════════════════════════════════════════════════
doc.add_page_break()
section_banner("PART 2 — Build a Website with Claude Code", "2D6A4F")

intro2 = doc.add_paragraph(
    "You don't need to know HTML, CSS, or JavaScript. Just describe what you want in plain "
    "English and Claude Code will write the code for you. Here's the full workflow."
)
intro2.paragraph_format.space_after = Pt(10)
for run in intro2.runs:
    run.font.size = Pt(10.5)
    run.font.color.rgb = GRAY

# Step 6
h3(6, "Open a New Project Folder", "IDE")
body("In your IDE, go to File → Open Folder (or Open…). Create a new empty folder on your "
     "computer (e.g. 'my-website') and open it. This is where your website files will live.")

divider()

# Step 7
h3(7, "Open the Claude Code Panel", "IDE")
body("Click the Claude Code icon in your IDE sidebar to open the chat panel. You'll see a "
     "text box where you can type instructions — just like texting or messaging.")

divider()

# Step 8
h3(8, "Describe Your Website", "IDE")
body("Type a description of the website you want. Be as specific as you like. For example:")
p = doc.add_paragraph(
    '"Build me a single-page campaign website for Jane Smith running for Belmont Town Meeting. '
    'Include a hero section with her name and tagline, an About section, a Priorities section '
    'with 3 issue cards, and a contact form. Use a navy and gold color scheme."'
)
p.paragraph_format.left_indent = Inches(0.4)
p.paragraph_format.space_after = Pt(6)
for run in p.runs:
    run.font.size = Pt(10)
    run.font.color.rgb = GRAY
    run.italic = True
body("Claude Code will generate all the files (index.html, style.css, main.js) automatically.")

divider()

# Step 9
h3(9, "Preview Your Site in the Browser", "IDE")
body("Once the files are created, right-click on index.html in the IDE file explorer and "
     "choose Open with Live Server or Open in Browser. Your website will open in your default "
     "browser so you can see exactly how it looks.")

divider()

# Step 10
h3(10, "Refine with Follow-Up Instructions", "IDE")
body("Don't like something? Just tell Claude Code what to change — no need to touch the code yourself. Examples:")
bullet("\"Change the background color to dark navy.\"")
bullet("\"Make the headline font larger and use Playfair Display.\"")
bullet("\"Add a fourth priority card about housing.\"")
bullet("\"Make the layout work better on mobile phones.\"")
body("Keep iterating until the site looks exactly the way you want.")

# ══════════════════════════════════════════════════════════════════
# PART 3: DEPLOY (PUBLISH LIVE)
# ══════════════════════════════════════════════════════════════════
doc.add_page_break()
section_banner("PART 3 — Deploy Your Website (Publish It Live)", "C8963E", NAVY)

intro3 = doc.add_paragraph(
    "Once your site looks good, it's time to put it on the internet so anyone can visit it. "
    "The options below are all free or very low cost, and require no command line."
)
intro3.paragraph_format.space_after = Pt(10)
for run in intro3.runs:
    run.font.size = Pt(10.5)
    run.font.color.rgb = GRAY

# ── Option A: GitHub Pages
h2("Option A — GitHub Pages (Free, Recommended)")

h3("A1", "Create a Free GitHub Account", "Web")
body("Go to github.com and sign up for a free account. GitHub is where your website files "
     "will be stored and served from.")
link_line("github.com", "https://github.com")

divider()

h3("A2", "Create a New Repository", "Web")
body("Once logged in, click the green New button (or the + icon at the top right). "
     "Name your repository — e.g. my-campaign-site. Set it to Public. "
     "Click Create repository.")

divider()

h3("A3", "Upload Your Files", "Web")
body("On your new repository page, click Add file → Upload files. "
     "Drag and drop your index.html, style.css, main.js, and any images into the upload area. "
     "Scroll down and click Commit changes.")

divider()

h3("A4", "Enable GitHub Pages", "Web")
body("Go to your repository Settings (tab at the top). In the left sidebar, click Pages. "
     "Under Source, select Deploy from a branch. Choose main and click Save. "
     "GitHub will give you a free URL like: https://yourusername.github.io/my-campaign-site")
body("Your site is now live. Share that link anywhere — social media, email, flyers.")

divider()

# ── Option B: Netlify
h2("Option B — Netlify Drop (Fastest, No Account Required)")

h3("B1", "Go to Netlify Drop", "Web")
body("Open your browser and go to app.netlify.com/drop. No account needed to try it.")
link_line("Netlify Drop", "https://app.netlify.com/drop")

divider()

h3("B2", "Drag & Drop Your Folder", "Web")
body("Drag your entire project folder (the one containing index.html, style.css, etc.) "
     "directly onto the Netlify Drop page. It will upload automatically and give you a "
     "live URL within seconds — e.g. https://random-name.netlify.app")

divider()

h3("B3", "Claim Your Site (Optional)", "Web")
body("Create a free Netlify account to keep the site permanently, rename the URL, "
     "and connect a custom domain name (e.g. meenalbagla.com) later.")

divider()

# ── Custom Domain
h2("Adding a Custom Domain Name (Optional)")
body("To use a custom URL like yourname.com or yourcampaign.com:")
bullet("Purchase a domain from Namecheap, Google Domains, or Squarespace Domains (~$12–15/year).")
bullet("In your GitHub Pages or Netlify settings, find the Custom Domain field and enter your domain.")
bullet("In your domain registrar's DNS settings, point the domain to GitHub or Netlify "
       "using the nameservers or CNAME/A records they provide. Both services have step-by-step guides.")
bullet("DNS changes take 15 minutes to 24 hours to fully propagate.")

# ── Form submissions
h2("Receiving Form Submissions (Optional)")
body("The contact form on your site can be connected to a free service so you receive emails "
     "when someone submits it:")
bullet("Go to formspree.io and create a free account.", "Formspree:")
bullet("Create a new form — Formspree gives you a unique endpoint URL.", "Create form:")
bullet('In your index.html, find the <form> tag and set action="https://formspree.io/f/YOUR_ID".', "Connect:")
bullet("Test by submitting the form. Responses will arrive in your email inbox.", "Test:")
link_line("formspree.io", "https://formspree.io")

# ══════════════════════════════════════════════════════════════════
# QUICK REFERENCE
# ══════════════════════════════════════════════════════════════════
doc.add_page_break()
section_banner("Quick Reference — All Tools & Links", "1B3A6B")

refs = [
    ("Claude Pro subscription",  "claude.ai",                   "https://claude.ai"),
    ("Claude Desktop download",  "claude.ai/download",          "https://claude.ai/download"),
    ("Cursor IDE download",       "cursor.com",                  "https://cursor.com"),
    ("GitHub (file hosting)",     "github.com",                  "https://github.com"),
    ("Netlify Drop (instant deploy)", "app.netlify.com/drop",   "https://app.netlify.com/drop"),
    ("Formspree (form emails)",   "formspree.io",                "https://formspree.io"),
    ("Namecheap (domains)",       "namecheap.com",               "https://namecheap.com"),
]

for label, display, url in refs:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    add_run(p, f"{label:<36}", bold=True, color=NAVY, size=10.5)
    add_run(p, f"  {display}", color=GRAY, size=10, italic=True)

divider()

footer_p = doc.add_paragraph()
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(footer_p,
        "Built with Claude Code · Anthropic  |  anthropic.com",
        color=GRAY, size=9, italic=True)

# ── Save ──────────────────────────────────────────────────────────
doc.save("/home/user/TechBit/Claude_Code_Guide.docx")
print("Saved: Claude_Code_Guide.docx")

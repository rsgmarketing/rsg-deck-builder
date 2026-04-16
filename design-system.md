# Presentation Design System

## Brand Selection

Ask the user which brand before starting. This determines colors, fonts, and logo selection.

| Brand | When to Use |
|---|---|
| **Norlake** | Walk-ins, Kold Locker, Fast-Trak, FineLine, Capsule Pak, Split-Pak, LogiTemp |
| **Master-Bilt** | Walk-ins, Endless merchandisers, ice cream cabinets, DC/DD/FLR series, Ready-Bilt |
| **Norlake Scientific** | Lab, scientific, vaccine, pharma refrigeration |
| **RSG Corporate** | Cross-brand, corporate strategy, investor, multi-brand |

If the content mentions products from both Norlake and Master-Bilt, use RSG Corporate branding.

---

## Color Palettes

### RSG Corporate (Default)

```python
# Primary
NAVY = RGBColor(0x00, 0x28, 0x57)       # #002857 — backgrounds, headers
DARK_NAVY = RGBColor(0x00, 0x15, 0x32)   # #001532 — darker panels
ACCENT_BLUE = RGBColor(0x2B, 0x7C, 0xCC) # #2B7CCC — accents, links, callouts
WHITE = RGBColor(0xFF, 0xFF, 0xFF)        # #FFFFFF — text on dark, backgrounds
LIGHT_GRAY = RGBColor(0xF0, 0xF1, 0xED)  # #F0F1ED — light slide backgrounds
BODY_GRAY = RGBColor(0x2C, 0x3E, 0x50)   # #2C3E50 — body text on light backgrounds

# Semantic
GREEN = RGBColor(0x10, 0xB9, 0x81)       # #10B981 — success, positive, ready
ORANGE = RGBColor(0xF5, 0x9E, 0x0B)      # #F59E0B — warning, caution, in-progress
RED = RGBColor(0xEF, 0x44, 0x44)          # #EF4444 — urgent, problem, risk
GOLD = RGBColor(0xD4, 0xA8, 0x43)         # #D4A843 — premium accent, dividers

# Tints (for card and panel backgrounds)
LIGHT_BLUE_BG = RGBColor(0xEB, 0xF5, 0xFF)  # #EBF5FF — info cards
GREEN_BG = RGBColor(0xEC, 0xFD, 0xF5)       # #ECFDF5 — success cards
RED_BG = RGBColor(0xFE, 0xF2, 0xF2)         # #FEF2F2 — warning cards
ORANGE_BG = RGBColor(0xFF, 0xFB, 0xEB)      # #FFFBEB — caution cards
```

### Norlake

Same palette as RSG Corporate. Norlake navy IS the RSG navy (#002857 / PMS 295).

### Master-Bilt

```python
# Master-Bilt uses a warmer palette
NAVY = RGBColor(0x1A, 0x1A, 0x2E)        # #1A1A2E — darker, warmer navy
ACCENT_BLUE = RGBColor(0x00, 0x6B, 0xB6)  # #006BB6 — Master-Bilt blue
# All other colors same as RSG Corporate
```

### Norlake Scientific

```python
# Scientific uses a cooler, clinical palette
NAVY = RGBColor(0x1B, 0x2A, 0x4A)        # #1B2A4A — scientific navy
ACCENT_BLUE = RGBColor(0x00, 0x7B, 0xC0)  # #007BC0 — clinical blue
# All other colors same as RSG Corporate
```

---

## Typography

### RSG Corporate / Master-Bilt / Norlake Scientific (Default)

Everyone has these fonts. Use these unless the user specifies Norlake branding.

```python
FONT_DISPLAY = 'Bebas Neue'    # Headlines, titles, large numbers
FONT_BODY = 'Segoe UI'         # Body text, bullets, labels, captions
```

### Norlake

Requires custom font installation. Warn the user if they may not have these.

```python
FONT_DISPLAY = 'Teko'              # Headlines, titles, large numbers
FONT_BODY = 'Trade Gothic Next'    # Body text (fallback: 'Segoe UI')
```

**Font installation warning:** If Teko or Trade Gothic Next are not installed on the machine that opens the PPTX, PowerPoint will substitute Arial/Calibri and all spacing will break. When in doubt, use the RSG Corporate fonts (Bebas Neue / Segoe UI) — they are universally available.

### Font Size Rules

| Element | Display Font Size | Body Font Size |
|---|---|---|
| Slide title (BLUF headline) | 36-48 pt | — |
| Section opener title | 48-56 pt | — |
| Subtitle / tagline | — | 18-24 pt |
| Body text | — | 12-14 pt |
| Bullet points | — | 11-13 pt |
| Stat callout (large number) | 56-72 pt | — |
| Stat label | — | 10-12 pt |
| Footer text | — | 8-9 pt |
| Owner badge text | — | 8-9 pt |
| Table cell | — | 9-11 pt |

**Never use auto-shrink.** python-pptx does not render text — it cannot detect overflow. Set explicit sizes and test.

---

## Slide Dimensions

```python
from pptx.util import Inches
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
```

Always 16:9 widescreen. Never 4:3.

---

## Layout Rules

### The 6.5-Inch Rule

**All content must stay above y = 6.5 inches.** The footer zone occupies the bottom 1 inch (y = 6.5" to 7.5"). Content placed below 6.5" will collide with the footer or be cut off in presentation mode.

### Margins

```python
MARGIN_LEFT = Inches(0.75)
MARGIN_RIGHT = Inches(0.75)
MARGIN_TOP = Inches(0.6)
CONTENT_WIDTH = Inches(11.833)   # 13.333 - 0.75 - 0.75
```

### Logo Placement

- **Title slides:** Centered or top-left, width 3-4 inches
- **Content slides:** Top-right corner, width 2-2.5 inches
- **Footer:** Small logo (1.5" wide) bottom-right, paired with footer text bottom-left

### Footer Convention

Every slide (except title and section dividers) should have a footer:

```python
# Footer bar
add_rect(slide, x=0, y=Inches(6.8), w=Inches(13.333), h=Inches(0.7), fill=NAVY)
# Footer text
add_text(slide, "CONFIDENTIAL", x=Inches(0.75), y=Inches(6.95), size=Pt(8), color=WHITE)
# Footer logo
slide.shapes.add_picture("assets/logos/[brand]-white.png", x=Inches(10.5), y=Inches(6.85), width=Inches(2))
```

---

## Slide Patterns

### 1. Title Slide
- Full navy background
- Large display font title (centered or left-aligned)
- Subtitle in lighter color below
- Logo and date at bottom
- No footer bar — the slide IS the branding

### 2. Section Divider
- Full navy or dark background
- Large display text (40-56 pt) centered
- Accent blue bar below title (3" wide, 4px tall)
- Optional subtitle in muted color
- Logo at bottom

### 3. Content Slide (Default)
- White or light gray background
- BLUF headline in display font (navy, 36-48 pt) at top
- Body content area below headline
- Footer bar at bottom

### 4. Stats Slide
- 3-4 stat blocks in a row
- Large number in display font (56-72 pt, accent blue)
- Label below in body font (10-12 pt, body gray)
- Even horizontal spacing across slide width

### 5. Comparison Slide
- Side-by-side colored panels (2-3 columns)
- Each panel: colored header bar + white body
- Equal widths with 0.2" gap between
- Good for product tiers, before/after, option A vs. B

### 6. Tier Cards
- 2-3 cards across with distinct top border colors
- Card: light background, colored top bar (4-6px), bold title, body text
- Use for product lines, pricing tiers, service levels

### 7. Question/Action Slide
- Numbered items with owner badges (right-aligned colored pills)
- Alternating row backgrounds (white / light blue)
- Good for action items, accountability, decision tracking

### 8. Closer/CTA Slide
- Navy background
- Clear call to action in large display text
- Contact information
- Logo centered

---

## Component Patterns

### Accent Bar
```python
def add_accent_bar(slide, x, y, width=Inches(3)):
    """Blue accent rule for visual rhythm"""
    add_rect(slide, x=x, y=y, w=width, h=Inches(0.05), fill=ACCENT_BLUE)
```

### Stat Block
```python
def add_stat_block(slide, number, label, x, y, number_size=Pt(56)):
    """Large number + label underneath"""
    add_text(slide, number, x=x, y=y, font=FONT_DISPLAY, size=number_size, color=ACCENT_BLUE, bold=True)
    add_text(slide, label, x=x, y=y + Inches(0.7), font=FONT_BODY, size=Pt(11), color=BODY_GRAY)
```

### Owner Badge
```python
OWNER_COLORS = {
    'MANUFACTURING': (RGBColor(0x92, 0x40, 0x0E), RGBColor(0xFF, 0xED, 0xD5)),
    'SALES':         (RGBColor(0x16, 0x65, 0x34), RGBColor(0xDC, 0xFC, 0xE7)),
    'FINANCE':       (RGBColor(0x6B, 0x21, 0xA8), RGBColor(0xF3, 0xE8, 0xFF)),
    'PRODUCT':       (RGBColor(0x1E, 0x40, 0xAF), RGBColor(0xDB, 0xEA, 0xFE)),
    'ENGINEERING':   (RGBColor(0x0E, 0x70, 0x90), RGBColor(0xCC, 0xFB, 0xF1)),
    'MARKETING':     (RGBColor(0xBE, 0x18, 0x5D), RGBColor(0xFC, 0xE7, 0xF3)),
}

def add_owner_badge(slide, owner, x, y):
    """Color-coded department pill"""
    text_color, bg_color = OWNER_COLORS.get(owner.upper(), (BODY_GRAY, LIGHT_GRAY))
    add_rect(slide, x=x, y=y, w=Inches(1.4), h=Inches(0.25), fill=bg_color, radius=Inches(0.12))
    add_text(slide, owner.upper(), x=x + Inches(0.1), y=y + Inches(0.03), size=Pt(8), color=text_color, bold=True)
```

### Card Panel
```python
def add_card(slide, title, body, x, y, w, h, accent_color=ACCENT_BLUE):
    """Card with colored top border"""
    # Top accent bar
    add_rect(slide, x=x, y=y, w=w, h=Inches(0.06), fill=accent_color)
    # Card body
    add_rect(slide, x=x, y=y + Inches(0.06), w=w, h=h - Inches(0.06), fill=WHITE, border=LIGHT_GRAY)
    # Title
    add_text(slide, title, x=x + Inches(0.2), y=y + Inches(0.15), font=FONT_DISPLAY, size=Pt(20), color=NAVY, bold=True)
    # Body text
    add_multiline(slide, body, x=x + Inches(0.2), y=y + Inches(0.55), w=w - Inches(0.4), font=FONT_BODY, size=Pt(11), color=BODY_GRAY)
```

# Deck Building Guide

How to create professional RSG presentations using Claude.

---

## Getting Started

Open this folder in Claude Cowork (or share the folder with Claude in chat). Claude reads the CLAUDE.md file automatically and knows how to build decks.

**What you say:**
> I need a 12-slide deck about [topic] for [audience]. It's for the [brand] brand.

**What Claude does:**
1. Asks any clarifying questions
2. Writes BLUF headlines for every slide and shows you for approval
3. Generates a Python script that builds the .pptx
4. Runs the script and gives you the file
5. Iterates based on your feedback

**You never write code.** Describe what you want in plain English. Claude handles the rest.

---

## What You Need to Provide

| What | Example | Required? |
|---|---|---|
| Topic / content source | "Our Fast-Trak pricing strategy" or paste a document | Yes |
| Audience | "Senior leadership" / "dealer partners" / "franchisees" | Yes |
| Brand | Norlake, Master-Bilt, RSG Corporate, or Norlake Scientific | Yes |
| Purpose | "Get approval for the pilot" / "convince them to stock our line" | Yes |
| Slide count | "About 12 slides" | Optional (Claude defaults to 10-15) |
| Specific content | Meeting notes, strategy docs, data | Optional but helpful |

**Tip:** The more context you give upfront, the better the first draft. Paste meeting notes, strategy docs, or even a rough bullet-point outline directly into the conversation.

---

## The Iteration Loop

The first draft will be ~80% right. The remaining 20% is where this workflow shines — iteration is fast.

### How to Give Good Feedback

**Be specific.** Reference slide numbers and describe what you see vs. what you want.

| Less Effective | More Effective |
|---|---|
| "Fix the layout on slide 5" | "On slide 5, the three callout boxes overlap the footer. Move them up." |
| "The text is too small" | "Body text on slide 7 is hard to read. Make it 14pt." |
| "Add more detail" | "Add a new slide after slide 10 comparing the three options in a table." |
| "I don't like it" | "The headline on slide 3 is a topic label. Make it a takeaway: 'R-290 cuts energy costs 50%.'" |

**Batch your feedback.** List 3-5 changes in one message rather than asking one at a time. Claude fixes them all and regenerates the entire deck.

### Common Changes

- **"Change the headline on slide X"** — Claude rewrites and regenerates
- **"Add a slide about X after slide Y"** — Claude inserts it in the right place
- **"Switch from Norlake to RSG branding"** — Claude swaps colors, fonts, and logos globally
- **"Make the stats bigger"** — Claude adjusts font sizes
- **"The title wraps to two lines"** — Claude shortens the text or reduces the font
- **"Use a different product image"** — Claude picks from the available inventory

---

## Understanding the Output

Claude generates a .pptx file — a standard PowerPoint file you can:
- Open in Microsoft PowerPoint or Google Slides
- Edit manually (move shapes, change text, add content)
- Present directly
- Export to PDF for distribution

The deck is built from colored rectangles, text boxes, and images at exact coordinates. Everything is editable in PowerPoint.

---

## QA Checklist

After downloading the .pptx, open it in PowerPoint and check:

- [ ] **Text overflow** — Any text cut off at the bottom or right edge of a box?
- [ ] **Unexpected wrapping** — Are titles or labels breaking to two lines?
- [ ] **Footer boundary** — Any content in the bottom inch of the slide?
- [ ] **Element overlap** — Do any cards, text boxes, or images overlap?
- [ ] **Logo placement** — Correct corner, right size, not overlapping content?
- [ ] **Font rendering** — Do fonts look correct? (See font note below)
- [ ] **Color consistency** — Same colors across all slides?
- [ ] **The conference room test** — Can you read everything from 6 feet away?

### Font Note

The default fonts are **Bebas Neue** (headlines) and **Segoe UI** (body text). These should be available on most machines.

For Norlake-branded decks, Claude uses **Teko** and **Trade Gothic Next**. If these aren't installed on your computer, PowerPoint will substitute Arial/Calibri and the spacing will look wrong. Either:
- Install the correct fonts before opening the .pptx
- Ask Claude to use Bebas Neue / Segoe UI instead

---

## What's in This Folder

| File | What It Does |
|---|---|
| `CLAUDE.md` | The instructions Claude follows (you don't need to read this) |
| `voice-rules.md` | Brand voice rules, BLUF examples, banned phrases |
| `design-system.md` | Colors, fonts, slide patterns, layout rules |
| `image-inventory.md` | Available logos and product images with usage guidance |
| `deck-building-guide.md` | This file — how to use the system |
| `assets/` | Logos, product images, and backgrounds |
| `knowledge/` | Product specs, competitive intel, audience personas |
| `examples/` | Reference deck and generator script showing what "good" looks like |

---

## Tips

1. **Start with the headline outline.** Getting the slide titles right is 60% of the work. If the headlines tell the story on their own, the deck will be strong.

2. **One idea per slide.** If you're trying to say two things on one slide, ask Claude to split it.

3. **Don't fight the system.** If Claude's layout doesn't match what you imagined, describe what you want — don't try to manually edit 15 slides of precise coordinates in PowerPoint.

4. **Use the reference deck.** Open `examples/Fast-Trak-Strategy-Reference.pptx` to see the quality level and style you should expect. If your deck looks different, tell Claude to match it.

5. **Proof points win.** Every slide should have at least one specific number, date, or certification. "We're the best" loses to "660K sq ft, ISO certified, 2-day shipping."

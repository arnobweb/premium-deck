---
name: premium-pitch-deck
description: "Create premium, highly visual PowerPoint presentations for executive decisions, strategy, fundraising, sales, projects, research, education and institutional communication. Use adaptable branding, clear narrative, custom diagrams, evidence-based claims and speaker notes. Use when a polished or investor-ready deck is requested; not for routine plain slides or when a different mandatory template is specified."
---

# Premium Deck

Create editable 16:9 PowerPoint decks with python-pptx. No proprietary assets or other skills are required. Use supplied branding or a neutral visual identity; never invent a logo or imply endorsement. Keep slides concise, visually confident and evidence-led.

## Narrative modes

Choose a mode based on the audience and intended outcome. These are adaptable structures, not mandatory slide sequences.

| Mode | Narrative |
|---|---|
| Executive | Decision context, evidence, options, trade-offs, recommendation, decision required |
| Strategy | Current state, ambition, choices, rationale, roadmap, measures |
| Commercial | Customer needs, solution, differentiation, evidence, proposal, next step |
| Project | Purpose, scope, deliverables, owners, milestones, risks, commitments |
| Scientific | Question, methods, findings, limitations, implications, references |
| Educational | Learning objectives, concepts, examples, practice, recap |
| Institutional | Mission, stakeholders, activities, evidence of impact, next steps |
| Investor | Investment thesis, opportunity, evidence, financials, funding ask |

Do not force an investment ask or a persuasive framing on educational or scientific material. Preserve uncertainty and distinguish evidence from interpretation. Use the source and requested audience to select layouts, length and level of detail.

## Investor narrative

For fundraising, select the relevant slides from this sequence, adapting to stage and audience: company and investment thesis; customer problem; solution and product; why now; market opportunity (TAM/SAM/SOM with definitions and sources where available); business model and pricing; traction and validation; competition and differentiation; go-to-market; team; financial performance and projections; funding ask, use of funds and milestones; closing and contact. Use an appendix for detailed assumptions and diligence. For non-investor decks, use the relevant narrative mode above.

Distinguish actual results from forecasts, targets and assumptions. Label dates, currencies, units and sources. Never fabricate traction, customers, market sizing, financials, team credentials or funding terms. Where evidence is absent, omit the claim or mark it as to be confirmed in a working draft; disclose gaps before external sharing. Do not present a conceptual diagram as quantitative data. Use proportionate charts for real metrics and avoid misleading scales. Adapt the ask to the actual decision; do not assume a funding amount or instrument.

## Workflow

1. **Read the source first** (PDF, note, text). Extract the story: problem, insight, solution, how it works, examples, value, risks, ask. Use `pdftotext -layout` for PDFs. Never invent figures: if budget, ROI, market size or dates are missing from the source, leave them out and say so in the final message.
2. **Ask 2-3 short questions before building** (using the available interaction mechanism): who the audience is, what decision or funding is sought, and whether brand assets and verified figures exist. Ask only for missing essential information. Confirm permission before using confidential material externally. Do not impose a confidentiality label.
3. **Propose the narrative** as a list of approximately 10-14 slide headlines, adjusted to the audience and available evidence, then build.
4. **Build** with the bundled `deck_lib.py` helper (import it, add your slides, save). Write speaker notes (2-4 sentences) for every slide.
5. **Visual QA loop** (required when rendering tools are available; otherwise disclose that visual QA could not be completed): `soffice --headless --convert-to pdf deck.pptx`, `pdftoppm -r 80 -png`, build contact sheets of 4-6 slides with PIL, look at them, fix overlaps, re-render. Check especially: titles wrapping to 3 lines, captions touching shapes, text clipped inside shapes.
6. **Deliver**: copy the .pptx to the user's workspace folder, present it, and in 3-4 lines state what was condensed or omitted and what is not in the source (e.g. no budget). Remind that the check was a LibreOffice render, not PowerPoint.

Timeouts: never use recursive `glob("/**")`; use only explicitly supplied logo paths. Run soffice with `timeout 90`.

## Design system

- Colors: navy `#072f3f`; navy-2 `#0c455a` (dark slides use a 35 degree gradient navy to navy-2); accent blue `#77a7cc`; teal `#3a7a8f`; blue tints `#9fbede`, `#bfd2e2`, `#dde8ef`; light background `#f4f6f8`; text `#333333`; grey `#6b6e73`; faint rings `#154a5e`. Use a supplied brand palette when available; otherwise use this neutral default.
- Font: Arial by default; use a supplied brand font if installed. Preheader 9 pt bold caps with letter spacing (spc 250). Titles 24-30 pt bold (full sentences, max 2 lines, 26 pt default; use 30 pt for short punchy titles). Body 10-13 pt. Minimum 9 pt except footers. No em dashes. Short phrases, max about 25 words per card.
- Rhythm: cover and closing dark. Alternate dark (opportunity, how it works, value, pilot, ask) and light (challenge, solution, use cases, governance, measures, risks). Never three of the same in a row if avoidable.
- Fixed elements (done by `slide()`): bottom blue band 0.125 in, optional logo bottom-left (contrast-appropriate, scaled by height only), footer with optional organization name and slide number bottom-right, decorative rings on dark slides, margins 0.6 in, content area y 1.35 to 4.9.
- Why the restraint: the design should convey quiet confidence. Use bold for hierarchy but avoid heavy decoration, drop shadows or clip art. Use shapes and numerals instead of icons or emoji (glyph support varies).

## Slide archetypes (pick the visual that matches the idea)

| Idea | Visual | Helper |
|---|---|---|
| Cover | dark, big two-line title, ghost acronym, ask line | `cover()` |
| Problem / gap | light, timeline with dotted gap + 3 white cards | `timeline_gap()`, `card_light()` |
| Why now / transformation | dark, source box to central node to 5 output pills | `fan_out()` |
| Principles (4) | light, colored header blocks with big numerals | `principle_cards()` |
| Process (4 steps) | dark, circles on a line + outcome strip | `numbered_flow()`, `strip()` |
| Examples (3) | light, 3 cards: colored header with persona letter, 3 bullets, "the shift" footer | `circ()`, `rect()`, `tb()` |
| Value / scaling | dark, ascending bars, last bar white | `stairs()` |
| Trust / governance | light, 3x2 white tiles with check circle + navy conclusion strip | `circ()`, `strip()` |
| Fit with existing programs | light, 4 rounded boxes with arrows, core box navy | `rect(...ROUNDED_RECTANGLE)`, `arrow()` |
| Pilot / roadmap | dark, giant number (88 pt) left + 3 chevrons + bullets + scope strip | `chevrons()` |
| Measures | light, navy question panel left + KPI rows with QUANT/QUAL pills | `pill()` |
| Risks / assumptions | light, 5 cards + navy "to be confirmed in..." tile | `card_light()` |
| The ask | dark, 3 step cards with numbered circles + attendee pills | `circ()`, `pill()` |
| Closing | dark, "Thank you", tagline, name and contact | `closing()` |

Adapt the sequence to the source; drop or add archetypes as needed. Layout coordinates in the helpers are tuned: reuse them rather than improvising, and keep 0.2 in gaps between cards and arrows centered in the gap.

## Content rules

- Title = the message of the slide as a sentence, not a label. The preheader carries the section label.
- One message per slide. Move detail to speaker notes.
- Use the source’s own terminology; define unfamiliar acronyms for external audiences.
- Keep claims traceable to the source. Mark assumptions as such. If the audience is "investors" but the source is internal and has no financials, say so rather than fabricating numbers.
- Follow the requested language, or the source language if unspecified.

## Helper library (`deck_lib.py`)

```python
"""Premium pitch deck helpers. 10 x 5.625 in, python-pptx, Arial by default; use a supplied brand font if installed."""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as C
from pptx.enum.shapes import MSO_SHAPE as SH, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN as AL, MSO_ANCHOR as AN
from pptx.enum.dml import MSO_LINE

NAVY=C(0x07,0x2f,0x3f); NAVY2=C(0x0c,0x45,0x5a); BLUE=C(0x77,0xa7,0xcc); T1=C(0x9f,0xbe,0xde)
T2=C(0xbf,0xd2,0xe2); T3=C(0xdd,0xe8,0xef); DB=C(0x3a,0x7a,0x8f); TX=C(0x33,0x33,0x33)
GR=C(0x6b,0x6e,0x73); LIGHT=C(0xf4,0xf6,0xf8); W=C(255,255,255); RING=C(0x15,0x4a,0x5e)
F="Arial"; STATE={"n":0}

BRAND={"name":"", "classification":"", "logo_light":None, "logo_dark":None}

def configure_brand(name="", classification="", logo_light=None, logo_dark=None, palette=None, font="Arial"):
    """Optional branding. Palette maps existing color names to six-digit hex strings."""
    global F
    BRAND.update(name=name, classification=classification, logo_light=logo_light, logo_dark=logo_dark)
    F=font
    for key,value in (palette or {}).items():
        if key not in {"NAVY","NAVY2","BLUE","T1","T2","T3","DB","TX","GR","LIGHT","W","RING"}:
            raise ValueError(f"Unknown palette key: {key}")
        globals()[key]=C.from_string(value.lstrip("#"))

def start():
    STATE["n"]=0; prs=Presentation(); prs.slide_width=Inches(10); prs.slide_height=Inches(5.625); return prs

def rect(s,x,y,w,h,fill=None,shape=SH.RECTANGLE,line=None,lw=1,adj=None):
    r=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill is None: r.fill.background()
    else: r.fill.solid(); r.fill.fore_color.rgb=fill
    if line is None: r.line.fill.background()
    else: r.line.color.rgb=line; r.line.width=Pt(lw)
    r.shadow.inherit=False
    if adj is not None and len(r.adjustments)>0: r.adjustments[0]=adj
    return r

def txt(sh,paras,anchor=AN.TOP,m=.1,mt=.05):
    """paras: list of (text,size,bold,color,align[,letter_spacing[,space_before_pt]])"""
    tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=tf.margin_right=Inches(m); tf.margin_top=tf.margin_bottom=Inches(mt)
    for k,pa in enumerate(paras):
        t,sz,b,col,al=pa[:5]; sp=pa[5] if len(pa)>5 else 0; sb=pa[6] if len(pa)>6 else 0
        p=tf.paragraphs[0] if k==0 else tf.add_paragraph(); p.alignment=al
        if sb: p.space_before=Pt(sb)
        r=p.add_run(); r.text=t; f=r.font; f.size=Pt(sz); f.bold=b; f.name=F; f.color.rgb=col
        if sp: r._r.get_or_add_rPr().set("spc",str(sp))

def tb(s,x,y,w,h,paras,anchor=AN.TOP):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); txt(t,paras,anchor,0,0); return t

def line(s,x1,y1,x2,y2,col,w=1.25):
    c=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    c.line.color.rgb=col; c.line.width=Pt(w); return c

def circ(s,x,y,d,fill,label=None,size=14,col=None,line_=None):
    col=col or W
    o=rect(s,x,y,d,d,fill,SH.OVAL,line=line_)
    if label is not None: txt(o,[(label,size,True,col,AL.CENTER)],AN.MIDDLE,0,0)
    return o

def arrow(s,x,y,w=.22,h=.3,c=None): return rect(s,x,y,w,h,c or BLUE,SH.RIGHT_ARROW)

def pill(s,x,y,w,h,label,fill=None,col=None,size=10):
    fill=fill or W; col=col or NAVY
    p=rect(s,x,y,w,h,fill,SH.ROUNDED_RECTANGLE,adj=.5); txt(p,[(label,size,True,col,AL.CENTER)],AN.MIDDLE,.05,0); return p

def strip(s,y,text,fill=None,col=None,size=11,h=.5):
    fill=fill or NAVY; col=col or W
    r=rect(s,.6,y,8.8,h,fill); txt(r,[(text,size,True,col,AL.CENTER)],AN.MIDDLE,.1,0); return r

def logo(s,dark,x=.6,y=5.04,h=.32):
    f=BRAND["logo_dark" if dark else "logo_light"]
    if f:
        if not os.path.isfile(f): raise FileNotFoundError(f)
        s.shapes.add_picture(f,Inches(x),Inches(y),height=Inches(h))

def _dark_bg(s):
    r=rect(s,0,0,10,5.625,NAVY); r.fill.gradient(); r.fill.gradient_angle=35
    r.fill.gradient_stops[0].color.rgb=NAVY; r.fill.gradient_stops[1].color.rgb=NAVY2

def slide(prs,dark,pre,title,notes,tsize=26):
    """Content slide shell: background, rings (dark), band, preheader, title, logo, footer, notes."""
    s=prs.slides.add_slide(prs.slide_layouts[6]); STATE["n"]+=1
    if dark:
        _dark_bg(s); rect(s,6.9,-1.6,4.6,4.6,None,SH.OVAL,line=RING); rect(s,7.7,-.8,3.0,3.0,None,SH.OVAL,line=RING)
    else: rect(s,0,0,10,5.625,LIGHT)
    rect(s,0,5.5,10,.125,DB if dark else BLUE)
    tb(s,.6,.38,8,.25,[(pre,9,True,T1 if dark else BLUE,AL.LEFT,250)])
    if title: tb(s,.6,.62,8.6,1.0,[(title,tsize,True,W if dark else NAVY,AL.LEFT)])
    logo(s,dark)
    tb(s,6.2,5.14,3.2,.2,[(" · ".join(filter(None,[BRAND["name"], BRAND["classification"], str(STATE["n"])])),8,False,T1 if dark else GR,AL.RIGHT,100)])
    s.notes_slide.notes_text_frame.text=notes
    return s

def cover(prs,pre,line1,line2,tagline,ask,author,ghost="",notes=""):
    s=prs.slides.add_slide(prs.slide_layouts[6]); STATE["n"]+=1; _dark_bg(s)
    for d,x,y in [(6.2,5.2,-.2),(4.6,6.0,.6),(3.0,6.8,1.4)]: rect(s,x,y,d,d,None,SH.OVAL,line=RING)
    tb(s,6.3,1.2,3.6,1.6,[(ghost,90,True,RING,AL.RIGHT)])
    tb(s,.6,.55,7,.25,[(pre,9,True,T1,AL.LEFT,250)])
    tb(s,.6,1.45,6.2,1.9,[(line1,44,True,W,AL.LEFT),(line2,44,True,W,AL.LEFT)])
    rect(s,.6,3.3,.7,.04,BLUE)
    tb(s,.6,3.5,5.8,.8,[(tagline,16,False,T2,AL.LEFT)])
    tb(s,.6,4.35,6,.3,[(ask,11,True,W,AL.LEFT)]); tb(s,.6,4.62,6,.25,[(author,10,False,T1,AL.LEFT)])
    logo(s,True,y=5.0); rect(s,0,5.5,10,.125,DB); s.notes_slide.notes_text_frame.text=notes; return s

def closing(prs,title,sub,name,contact,notes=""):
    s=prs.slides.add_slide(prs.slide_layouts[6]); STATE["n"]+=1; _dark_bg(s)
    for d,x,y in [(6.2,5.2,-.2),(4.6,6.0,.6),(3.0,6.8,1.4)]: rect(s,x,y,d,d,None,SH.OVAL,line=RING)
    tb(s,.6,1.5,6.5,1.0,[(title,44,True,W,AL.LEFT)]); rect(s,.6,2.55,.7,.04,BLUE)
    tb(s,.6,2.8,6.5,.6,[(sub,16,False,T2,AL.LEFT)])
    tb(s,.6,3.7,6.5,.5,[(name,12,True,W,AL.LEFT),(contact,10,False,T1,AL.LEFT,0,2)])
    logo(s,True,y=5.0); rect(s,0,5.5,10,.125,DB); s.notes_slide.notes_text_frame.text=notes; return s

# ---- composite components ------------------------------------------------
def card_light(s,x,y,w,h,head,body,bar=None,num=None):
    """White card, left accent bar, optional number, bold head, small body."""
    rect(s,x,y,w,h,W); rect(s,x,y,.06,h,bar or BLUE)
    off=.2
    if num: tb(s,x+.2,y+.08,.6,.3,[(num,14,True,BLUE,AL.LEFT)]); off=.65
    tb(s,x+off,y+.1,w-off-.1,.3,[(head,12,True,NAVY,AL.LEFT)])
    tb(s,x+.2,y+.45,w-.35,h-.5,[(body,10,False,TX,AL.LEFT)])

def principle_cards(s,items,y=1.75,h=2.65):
    """items: [(num,head,body)] x4. Colored header block with big number, white body."""
    cols=[BLUE,DB,BLUE,NAVY]
    for i,(n,hd,b) in enumerate(items):
        x=.6+i*2.2; rect(s,x,y,2.05,h,W); rect(s,x,y,2.05,.9,cols[i%4])
        tb(s,x+.18,y+.1,1.7,.7,[(n,28,True,W,AL.LEFT)],AN.MIDDLE)
        tb(s,x+.18,y+1.05,1.75,.35,[(hd,15,True,NAVY,AL.LEFT)]); tb(s,x+.18,y+1.45,1.75,1.1,[(b,11,False,TX,AL.LEFT)])

def numbered_flow(s,steps,y=2.45):
    """Dark-slide 4-step flow: circles on a line, title and caption below. steps: [(n,head,body)]"""
    line(s,1.3,y,8.7,y,DB,2)
    for i,(n,h,b) in enumerate(steps):
        cx=1.3+i*2.47; circ(s,cx-.35,y-.35,.7,BLUE if i%2==0 else T1,n,20,NAVY)
        tb(s,cx-1.0,y+.55,2.0,.5,[(h,13,True,W,AL.CENTER)]); tb(s,cx-1.0,y+1.05,2.0,.9,[(b,10,False,T2,AL.CENTER)])

def fan_out(s,source_head,source_body,formats,node="CORE",source_label="INPUT"):
    """Dark slide: source box -> central node -> up to 5 format pills."""
    src=rect(s,.6,2.0,2.5,2.3,NAVY2,SH.ROUNDED_RECTANGLE,line=T1,adj=.06)
    txt(src,[(source_label,9,True,T1,AL.CENTER,200),(source_head,20,True,W,AL.CENTER,0,4),(source_body,10,False,T2,AL.CENTER,0,8)],AN.MIDDLE,.2)
    circ(s,4.0,2.7,.9,BLUE,node,12,NAVY); line(s,3.1,3.15,4.0,3.15,T1,1.5)
    for i,f in enumerate(formats):
        yy=1.7+i*.58; line(s,4.9,3.15,6.2,yy+.21,T1,1.0); pill(s,6.2,yy,3.2,.43,f,W if i%2==0 else T3,NAVY,12)

def stairs(s,steps):
    """Dark slide ascending bars. steps: [(label,head,body)] x4; last bar is white."""
    hs=[1.3,1.85,2.4,2.95]; fills=[DB,DB,BLUE,W]
    for i,(a,h,b) in enumerate(steps):
        x=.6+i*2.2; top=4.9-hs[i]; r=rect(s,x,top,2.05,hs[i],fills[i])
        c=NAVY if i>=2 else W; lc=NAVY2 if i==2 else (DB if i==3 else T2)
        txt(r,[(a,9,True,lc,AL.LEFT,150),(h,13,True,c,AL.LEFT,0,3),(b,10,False,c,AL.LEFT,0,3)],AN.TOP,.15,.12)

def chevrons(s,phases,x0=3.2,step=2.05,y=1.75):
    """Process chevrons. phases: [(label,head)]. Put bullets below each at x0+i*step+.1, y+.85."""
    cols=[DB,BLUE,T1]
    for i,(a,h) in enumerate(phases):
        c=rect(s,x0+i*step,y,2.1,.7,cols[i%3],SH.PENTAGON if i==0 else SH.CHEVRON,adj=.3)
        fc=W if i==0 else NAVY; txt(c,[(a,8,True,fc,AL.CENTER,150),(h,11,True,fc,AL.CENTER)],AN.MIDDLE,.3,0)

def timeline_gap(s,left,right,gap_label,y=2.55):
    """Light slide: two dots joined by a dotted line with a gap label. left/right: (LABEL, caption)"""
    rect(s,.9,y-.02,8.2,.04,T2); circ(s,.75,y-.2,.4,NAVY); circ(s,8.85,y-.2,.4,DB)
    for cx in [2.9+i*.6 for i in range(8)]: circ(s,cx,y-.05,.1,T2)
    tb(s,.6,y-.73,3,.3,[(left[0],9,True,NAVY,AL.LEFT,150)]); tb(s,.6,y+.45,2.6,.5,[(left[1],10,False,GR,AL.LEFT)])
    tb(s,6.6,y-.73,3,.3,[(right[0],9,True,DB,AL.RIGHT,150)]); tb(s,6.6,y+.45,3,.5,[(right[1],10,False,GR,AL.RIGHT)])
    p=rect(s,3.7,y-.5,2.6,.3,W,SH.ROUNDED_RECTANGLE,line=T1,adj=.5); txt(p,[(gap_label,10,True,DB,AL.CENTER)],AN.MIDDLE,0,0)
```

## Minimal usage

```python
import sys; sys.path.insert(0,".")
from deck_lib import *
prs=start()
cover(prs,"INVESTOR PRESENTATION","Company","Name",
      "One-line value proposition","Funding objective, if confirmed","Founder · Team",notes="...")
s=slide(prs,False,"THE CHALLENGE","Customers face a costly unresolved problem.","notes...",tsize=24)
timeline_gap(s,("CURRENT STATE","Existing workflow"),("DESIRED STATE","Customer outcome"),"Unresolved gap")
# ... more slides, then:
closing(prs,"Thank you","Tagline","Name","Team · email")
prs.save("deck.pptx")
```

Cards that contain text must be sized for it: about 40 characters per line at 10 pt in a 2.8 in card. If a render shows overflow, shorten the text before shrinking the font.


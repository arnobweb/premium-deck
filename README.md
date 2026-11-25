# Premium Deck

**Turn complex information into clear, compelling and decision-oriented presentations.**

Premium Deck is an AI-powered presentation framework combining strategic storytelling, visual design, information integrity and automated PowerPoint production. It provides a reusable methodology to transform raw information into structured, meaningful and visually consistent presentations.

## Who is it for?

| User | Applications |
|---|---|
| Entrepreneurs | Investor pitches, fundraising, startup presentations |
| Consultants | Strategic recommendations, audit findings |
| Managers | Business reviews, executive presentations, strategy |
| Sales professionals | Commercial proposals, customer presentations |
| Project managers | Project kickoffs, funding requests |
| Researchers | Research findings, grant applications |
| Independent professionals | Service presentations, partnership proposals |
| Nonprofit organizations | Fundraising, institutional communication |
| Communication agencies | Professional presentation production |

Especially useful for teams without dedicated designers or communication specialists.

## Why use a skill instead of a simple prompt?

The framework is designed to establish a repeatable production standard rather than requiring users to restate their expectations each time.

| Standard prompting may involve | Premium Deck aims to provide |
|---|---|
| Repeated instructions | Reusable methodology |
| Variable design | Consistent design principles |
| Text-heavy slides | Visual information hierarchy |
| Collection of slides | Structured narrative |
| Approximate brand alignment | Customizable visual identity |
| Difficult-to-reproduce results | Common production standard |

## Four complementary capabilities

- **Strategic storytelling:** adapt the argument to the audience, message and intended outcome.
- **Visual design:** select diagrams, comparisons, timelines and layouts that convey meaning.
- **Information integrity:** distinguish facts, estimates, assumptions and recommendations; never fabricate evidence.
- **Technical production:** create editable PowerPoint slides using a reusable Python library, with speaker notes and customizable branding.

## Design philosophy

Audience before content. Narrative before design. Meaning before decoration. Evidence before persuasion. Clarity before complexity. Consistency without uniformity.

Every slide should have a purpose, every visual should convey meaning, and every presentation should support a clear objective.

## One framework, multiple applications

| Application | Objective |
|---|---|
| Investor | Secure investment |
| Executive | Support decisions |
| Strategy | Build alignment |
| Sales | Demonstrate value |
| Project | Coordinate execution |
| Scientific | Communicate evidence |
| Educational | Transfer knowledge |
| Institutional | Communicate purpose and impact |

These are intended applications, not independently validated workflows. The bundled example demonstrates an investor deck; other applications require testing. The framework supports human judgment rather than replacing it.

**The objective is not simply to produce more slides, but to communicate better.**

## Features

- Editable 16:9 PowerPoint output, reusable layouts and speaker notes.
- Customizable colors, fonts, optional logos and confidentiality labels.
- Multilingual output, with English skill instructions.
- Visual quality checks when rendering tools are available.

## Install in Claude

Download `premium-pitch-deck-skill.zip` from this repository. In Claude, if custom Skills are enabled, open Settings / Customize > Skills and upload the ZIP. The archive contains a skill folder with `SKILL.md` and its helper. Interface labels and availability may vary by account. For Claude Code, copy `skills/premium-pitch-deck/` into `.claude/skills/` in your project (or `~/.claude/skills/` for personal use). The repository also contains plugin metadata for compatible Claude Code environments.

## Example prompt

> Create an investor-ready pitch deck from my attached business plan and financial model. Use our branding, distinguish actual results from projections, highlight our competitive advantage, and clearly explain the funding ask and use of funds. Include speaker notes and deliver an editable PowerPoint. Ask me for missing essential information rather than inventing it.

For other audiences, try: “Create an executive decision deck from these documents. Explain the options, trade-offs, evidence and recommended next steps. Use editable visuals and speaker notes.”

## Run locally without Claude

Requires Python 3.10+.

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python examples/build_example.py
```

The example writes `examples/example.pptx`. It is a fictional, illustrative presentation, not a real investment opportunity. The helper is a slide-building library, not an automatic narrative generator.

For visual QA, install LibreOffice and Poppler separately. Render the PowerPoint to PDF and inspect every slide. Rendering may differ from Microsoft PowerPoint.

## Branding

Call `configure_brand(name="Your company", font="Arial", palette={"NAVY":"#102A43", "BLUE":"#60A5FA"})` before building slides. Optional `logo_light`, `logo_dark` and `classification` arguments are supported. Use the module namespace when accessing customized palette constants. Check contrast and logo dimensions after customization.

## Repository layout

- `skills/premium-pitch-deck/SKILL.md`: AI instructions and reference implementation
- `skills/premium-pitch-deck/deck_lib.py`: reusable PowerPoint helper
- `examples/`: runnable example and editable output
- `.claude-plugin/plugin.json`: plugin metadata
- `premium-pitch-deck-skill.zip`: standalone upload package

## License

Premium Deck is licensed under the [MIT License](LICENSE).

Copyright (c) 2026 Arnaud Brihay.

from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "skills/premium-pitch-deck"))
import deck_lib as d

d.configure_brand(name="Example venture", classification="Illustrative")
p = d.start()
d.cover(p, "ILLUSTRATIVE INVESTOR DECK", "Example", "Venture", "A clearer way to manage complex workflows", "Concept only | funding not specified", "Example team", notes="Fictional example for demonstrating layout. No investment claims are made.")
s = d.slide(p, False, "THE CHALLENGE", "Fragmented workflows create friction", "Illustrative problem statement; validate with customer evidence.")
d.timeline_gap(s, ("TODAY", "Disconnected tools"), ("DESIRED STATE", "Coordinated workflow"), "Workflow gap")
s = d.slide(p, True, "THE SOLUTION", "A single workflow connects the steps", "Conceptual process, not validated product performance.")
d.strip(s, 3.8, "Capture  |  Organize  |  Review  |  Act")
d.closing(p, "Thank you", "Next step: validate the concept", "Example team", "Illustrative only", notes="Replace with verified company information before external use.")
p.save(str(Path(__file__).with_name("example.pptx")))

# AI Business Strategy Agent

A Claude Skill that turns one question, **"Give me [company]'s strategy"**, into a researched, consulting-style strategy deck.

The agent researches the company from public filings and news, then runs the **Define, Create, Capture, Deliver (DCCD)** framework in order, so each stage feeds the next. A Python pipeline renders the result as a 9-slide PowerPoint deck. The agent also answers follow-up questions concisely and defends its conclusions when challenged.

Built for 46-882 Business Strategy Fundamentals (Carnegie Mellon, Tepper School of Business, MSBA), where it was tested live on an unseen company.

![DCCD statements slide](docs/statements.png)

## What it produces

| # | Slide | What it answers |
|---|---|---|
| 1 | Title | |
| 2 | DCCD statements | The strategy in four statements, plus the competitive advantage |
| 3 | DCCD chain | How each stage feeds the next |
| 4 | Define | Who is the customer, what is the offer, what problem it solves |
| 5 | Create (map) | Willingness to pay vs. cost, plus a competitive positioning map |
| 6 | Create (VRI) | Are the resources Valuable, Rare and Inimitable? |
| 7 | Capture | Six forces: how much value stays with the firm |
| 8 | Deliver (activities) | Key activities, each linked to Define, Create and Capture |
| 9 | Deliver (fit) | Reinforcing choices, trade-offs, misaligned activities, judgment |

| Positioning map | Six forces |
|---|---|
| ![Create slide](docs/create_map.png) | ![Capture slide](docs/capture.png) |

| DCCD chain | Activity fit |
|---|---|
| ![Chain slide](docs/chain.png) | ![Deliver slide](docs/deliver_fit.png) |

## How it works

1. **Research:** the agent pulls the 8 to 15 facts the argument rests on (10-K, earnings, investor materials), each with a source. It never invents a number.
2. **Reason:** it runs Define, Create, Capture and Deliver in order, then checks they are consistent (for example, that Deliver defends against the strongest force named in Capture).
3. **Spec:** it writes the analysis into a JSON spec ([schema in SKILL.md](dccd-strategy-agent/SKILL.md)).
4. **Render:** `build_deck.py` turns the spec into the deck: a fixed type scale, a positioning map with a concave competitive frontier, styled tables, highlighted key rows, and overflow warnings if any text would not fit.

Besides decks, the skill has four other modes: a single DCCD statement, CEO-style Q&A, defending a claim under pushback, and explaining its own design choices.

## Repository layout

```
dccd-strategy-agent/          the Claude Skill
  SKILL.md                    instructions: modes, research protocol, DCCD logic, deck rules
  scripts/build_deck.py       JSON spec -> .pptx deck
  assets/template.pptx        slide master
  examples/                   specs and finished decks: Nike, Capital One, Costco, Tesla, Duolingo
docs/                         slide previews
```

## Use it

**As a Claude Skill (claude.ai):** zip the `dccd-strategy-agent` folder, then upload it in Settings > Capabilities > Skills (code execution and web search on). Ask: *"Give me Duolingo's strategy."*

**In Claude Code:** copy `dccd-strategy-agent` into `~/.claude/skills/`.

**Render a deck from a spec directly:**

```bash
pip install -r requirements.txt
cd dccd-strategy-agent
python scripts/build_deck.py examples/duolingo_spec.json Duolingo_DCCD.pptx assets/template.pptx
```

## Team

Developed by group "9 is Fine!" (Zuhayr Rahman Chowdhury and teammates), Tepper MSBA, Fall 2026.

## License

MIT for the code. The analyses in `examples/` are illustrative course work, not investment advice.

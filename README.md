# VC Deal Memo Skill

A personal Codex skill for drafting and polishing finance-grade VC deal memos, LP/syndicate memos, portfolio follow-on memos, sector coverage, diligence summaries, and IC-style recommendations.

The skill is designed for memos that need to sound like a strong venture investor wrote them: sponsor-grade, concise, valuation-aware, and grounded in verified company materials.

## Install With Codex Skill Installer

In Codex, ask:

```text
Use $skill-installer to install https://github.com/AmoghReddy45/vc-deal-memo-skill/tree/main/vc-deal-memo
```

Then restart Codex to pick up the new skill.

If installing from a terminal, this verified command uses git directly:

```bash
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo AmoghReddy45/vc-deal-memo-skill \
  --path vc-deal-memo \
  --method git
```

## Manual Install

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/AmoghReddy45/vc-deal-memo-skill.git
cp -R vc-deal-memo-skill/vc-deal-memo ~/.codex/skills/
```

Then restart Codex.

## Usage

```text
Use $vc-deal-memo to draft a Syndicate memo from these company materials and round terms.
```

```text
Use $vc-deal-memo to tighten this memo so it sounds more sponsor-grade and less like a generic analyst note.
```

## What It Covers

- Syndicate / LP memo structure
- Internal investment memo structure
- Portfolio follow-on framing
- Investment thesis bullets
- Valuation and comparable-company checks
- ARR vs run-rate vs revenue language
- Risk and watch-item framing
- DOCX / PDF formatting standards

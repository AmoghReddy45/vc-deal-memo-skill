---
name: vc-deal-memo
description: Use when drafting, revising, or polishing VC investment memos, syndicate/LP deal memos, portfolio follow-on memos, sector coverage, diligence summaries, round overviews, or IC-style recommendations. Applies when the user wants finance-grade structure, sponsor voice, valuation/comps logic, risk framing, round details, DOCX/PDF memo formatting, or memo copy that sounds like a strong venture investor rather than a generic analyst.
---

# VC Deal Memo

## Overview

Create investment memos that are clear enough for LPs, sharp enough for partners, and polished enough to send. The default style is sponsor-grade: bullish where earned, institutionally balanced, concise, and grounded in verified company materials, public sources, or clearly labeled assumptions.

## First Classify the Memo

Before drafting, identify the mode:

- **Syndicate / LP memo:** clean sponsor voice, investable framing, limited jargon, key facts, watch-items, round details, subscription-safe language.
- **Internal investment memo:** deeper judgment, diligence gates, decision math, valuation sensitivity, bear case, and unresolved questions.
- **Portfolio follow-on memo:** what changed since first investment, current ownership/pro-rata context, whether price still works, follow-on sizing, and fund-fit.
- **Sector coverage memo:** market map, competitors, pricing comps, product vectors, and why this company is advantaged or exposed.

If the user does not specify, infer from the audience and output format.

## Core Workflow

1. **Gather context:** company materials, financials, round terms, investment history, ownership basis, customer/traction data, product surface, team notes, comps, and disclosure constraints.
2. **Separate public vs confidential:** never include sensitive customer names, financials, cap table, or process details unless the user says they are approved for the memo.
3. **Build the spine:** recommendation / why now, what changed, why price works, key facts, thesis, comps, traction, market, team, risks, conclusion, round details.
4. **Write sponsor-grade copy:** direct, concise, finance-native, and not over-hedged. Use "company materials show/report/list" when relying on company-provided data.
5. **Validate finance logic:** ARR vs run-rate, gross vs net revenue, valuation cap vs priced valuation, post-money math, ownership basis, dilution, pro-rata, multiples, comp cleanliness.
6. **Polish the artifact:** tables where they improve skimmability, compact bullets, section bars, consistent headers, no orphan sections, and visual QA for DOCX/PDF outputs.
7. **Use deterministic formatting when needed:** for polished DOCX output, prefer `scripts/house_memo_builder.py` over manually recreating layout. Convert the memo into a JSON outline and run the builder.

## Reference Routing

Load only the references needed for the task:

- For section order and memo architecture, read `references/structure.md`.
- For language, tone, thesis bullets, sponsor voice, and non-generic writing craft, read `references/style.md`.
- For valuation, comps, ARR/run-rate, pro-rata, and diligence math, read `references/finance-checks.md`.
- For DOCX/PDF layout, tables, and presentation standards, read `references/formatting.md`.
- For deterministic house-style DOCX generation, read `references/builder.md` and use `scripts/house_memo_builder.py`.

## Non-Negotiables

- Do not invent signals, customer names, revenue, investor participation, or insider process details.
- Do not hide uncertainty; place it in the right container: "watch-item," "diligence question," "company-materials basis," or "to confirm in internal materials."
- Do not use generic VC filler such as "massive market," "AI tailwinds," "category-defining" unless the memo proves the claim.
- Do not write in generic AI prose. Prefer plain, finance-native sentences with specific nouns, clear verbs, and visible judgment.
- Do not tell the recipient firm to confirm facts they already know. Prefer "actual sizing depends on internal cap-table position" over "confirm current ownership."
- Do not overdo product description when the audience already knows the company. Prioritize why the business is now investable.
- For LP-facing memos, use "watch-items" more often than "bear case" unless the house style says otherwise.
- For DOCX/PDF deliverables, render and inspect the document before finalizing.
- If another model is struggling with formatting, do not keep adding prose instructions. Use the builder script or a template-based workflow.

# House Memo Builder

Use this when formatting quality matters. Do not ask the model to manually recreate the memo layout. Instead:

1. Draft the memo content.
2. Convert it into the JSON outline below.
3. Run `scripts/house_memo_builder.py`.
4. Open/render the DOCX and inspect the result.

## Command

```bash
python3 /path/to/vc-deal-memo/scripts/house_memo_builder.py memo_outline.json output.docx
```

If `python-docx` is missing:

```bash
python3 -m pip install python-docx
```

## JSON Shape

```json
{
  "title": "Company Syndicate Deal Memo",
  "subtitle": "Prepared for Example Syndicate | July 2026",
  "confidential": true,
  "footer": "Example Syndicate | Confidential company information",
  "meta": [
    {"label": "Opportunity", "value": "$40M SAFE at $4B valuation cap"},
    {"label": "Decision", "value": "Participate up to the full allocation"}
  ],
  "sections": [
    {
      "title": "Overview",
      "paragraphs": [
        "Company is raising [terms]. The case for participating is [reason]."
      ],
      "deal_table": {
        "rows": [
          ["Sector", "Data for AI / AI infrastructure"],
          ["Raise Terms", "$40M SAFE at a $4B valuation cap"],
          ["Syndicate Allocation", "$2,000,000"]
        ]
      }
    },
    {
      "title": "Our Investment Thesis",
      "bullets": [
        {"label": "Demand from the right buyers", "text": "Customer materials list [buyer categories]. These are the budgets that matter most."},
        {"label": "Price discipline", "text": "The round is not cheap. The price can work if current run-rate quality, customer breadth, and margin durability hold."}
      ]
    },
    {
      "title": "Comparable Companies and Transactions",
      "paragraphs": [
        "These comps are directionally useful, but not clean. Reported revenue may be gross billings, run-rate, or recognized revenue."
      ],
      "tables": [
        {
          "headers": ["Company / Transaction", "Valuation Context", "Revenue / Run-Rate Context", "Read-Through"],
          "rows": [
            ["Target", "$X valuation", "$Y revenue", "Pricing read for the target."],
            ["Comp A", "$X valuation", "$Y run-rate", "What this comp does and does not prove."]
          ]
        }
      ]
    }
  ]
}
```

## Supported Section Fields

- `title`: section-bar title.
- `paragraphs`: plain paragraphs. `**bold**` is supported.
- `deal_table`: two-column key-value table with house style.
- `bullets`: list of strings or `{ "label": "...", "text": "..." }`.
- `tables`: standard tables with `headers`, `rows`, optional `widths`, and optional `numeric_columns`.
- `callouts`: list of `{ "label": "...", "text": "..." }`.
- `page_break_before`: boolean.

## Style Behavior

The builder applies the house style automatically:

- Arial 10.2 pt body.
- Black section bars with white bold text.
- Light gray table headers.
- Thin 0.5 pt table borders.
- Content-sized columns.
- Special "The Deal" table with gray left/right columns and only perimeter border.
- Compact paragraph spacing.
- Em dashes and en dashes converted to hyphens.


# Related-party statement domain rules

## NON-NEGOTIABLE: quarters are calendar-year-to-date

Never implement a quarter as a standalone three-month block.

- 1Q = January through March
- 2Q = January through June
- 3Q = January through September
- 4Q = January through December

All period-dependent summary and detail output must consume the same ledger
already filtered by `Period`. In particular, 2Q detail includes Q1 transactions.
Do not globally replace `1Q` text in transaction descriptions: a real Q1
description is valid inside a Q2 YTD ledger.

Any code change touching period selection, ledger extraction, aggregation,
`39.1`, `39.2`, or statement generation must run the period, detail, template,
pipeline tests that exist at that point, plus the full test suite, before completion.

Never commit `_sample_input*`, `testrun/`, generated statements, or other
accounting data.

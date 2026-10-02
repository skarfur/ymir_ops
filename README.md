# ymir_ops

Tools that turn the club's Obsidian notes into a relational data set (a Google Sheet) that the
finance plan, calendar and later a website can read from.

```
Obsidian vault ──► extraction ──► data/*.yaml ──► tools/build_sheet.py ──► workbook ──► Google Sheet
(ymir_vault)       (parser + Claude)  (private)       (checks + builds)
```

This repository is public, so it holds only the schema and the code. The notes live in the private
`ymir_vault` repo, and the extracted data (names, prices, finances) is never committed here:
`data/` and `*.xlsx` are git-ignored.

## The data model

`schema/types.yaml` defines everything:

- **types**: one tab each (Principle, Goal, Priority, Decision, Task, Milestone, Division, Body,
  Function, Person, Organisation, Place, Programme, Asset, Price, Figure, Funding, Meeting, Insight)
- **common columns** on every tab: `status`, `basis` (notes / proposal / mixed), `conflict`,
  `source` (the vault note), `quote`, `origin` (extracted / human)
- **relations** for the Links tab (`serves`, `blocks`, `advances`, `priced_as`, `funds`, …), with
  the types each end may be

Adding a type, a field or a relation is a change to that file and a rebuild. Nothing else is
hard-coded to the current list.

## Building

```
pip install pyyaml openpyxl
python tools/parse_ops.py <vault_dir> <data_dir>      # Operations/ tasks, priorities, decisions
python tools/build_sheet.py <data_dir> <out.xlsx>     # validate and build the workbook
```

`parse_ops.py` reads the structured notes in the vault's `Operations/` folder (Tasks-plugin lines
and tables). Everything else is extracted by reading the notes and written as `curated_*.yaml` in
the data directory. Rows there with the same id as a parsed row are merged into it.

`build_sheet.py` refuses to build if an id is duplicated, a link points at a row that doesn't exist,
a relation isn't in the schema or joins the wrong types, or a row has a field the schema doesn't
define. In links, `~Task:some words` refers to the one Task whose name contains those words.

## Rules the data follows

- Every row names the note it came from. Contradictions between notes go in `conflict` and are
  collected on the Review tab, rather than one value being picked silently.
- `basis: proposal` marks things that exist only in the `Operations/` synthesis (rankings, owners,
  dates), so they read as proposals, not decisions.
- Rows a person has edited get `origin: human`; later extractions must not overwrite them.
- Children's names and personal details are left out, and reference material (coaching guides,
  grant-abstract analysis) is not extracted as data.

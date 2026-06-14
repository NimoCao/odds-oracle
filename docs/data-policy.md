# Data Policy

`odds-oracle` treats betting data as private by default.

## Public

Safe to commit:

- Source scripts in `scripts/`
- Project documentation
- Blank workbook templates in `templates/`
- Empty directory markers such as `data/.gitkeep`

## Private

Keep in `data/` only:

- Real betting ledgers
- Stake sizes, odds, settlement results, and profit/loss
- Filled prediction workbooks
- Scraped CSV/JSON snapshots
- Workbook backups and rendered previews
- Private scripts that contain concrete picks, fixtures, notes, or betting recommendations

Before each push, run:

```bash
git status --short
git check-ignore -v data/private-file-example.xlsx
```

If a private file appears as staged or untracked outside `data/`, move it into `data/` before committing.

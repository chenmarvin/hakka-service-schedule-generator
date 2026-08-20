# 客語堂服事季報表生成器 (Hakka Service Quarterly Schedule Generator)

A small Flask app that generates a quarterly Sunday service duty roster
(emcee / pianist / projectionist) for a church, subject to configurable
constraints (unavailable dates, pairing rules, per-person role caps, etc.),
and exports the result to `.xlsx`.

## Setup

```bash
pip install flask openpyxl
cp config.example.json config.json   # fill in your own roster & constraints
python3 app.py
```

Then open http://localhost:5000.

## Notes

- `config.json` and generated `.xlsx`/`.docx` schedules are git-ignored
  since they contain real names — copy `config.example.json` to get started.
- `generate_schedule.py` holds the scheduling/constraint-solving logic and
  the Excel export.
- `find_hakka_roles.py` is a helper script for extracting role/person data.

# Representative workbook output

Run the public scheduler and workbook builder from the repository root:

```bash
python gen7_engine.py
python build8_workbook_builder.py
```

The builder writes `Crew_Schedule_SAMPLE.xlsx` here. The generated workbook uses the repository's pseudonymous roster, representative demand, placeholder labor targets, undated Wednesday-through-Tuesday labels, and generalized decision notes.

The workbook is generated rather than committed so a stale binary cannot preserve details that have since been removed from the public source. Reviewers should treat the Python files and `schedule_data.py` as the public source of truth.

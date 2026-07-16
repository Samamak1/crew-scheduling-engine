# Restaurant Crew Scheduling Engine

> An AI-assisted workforce scheduling system translating 43-person availability, station qualifications, training gates, and coverage rules into a leadership-ready Excel workbook.

The original versions of this tooling were used by Sama Mushtaq, then the restaurant manager responsible for the schedule, to produce posted weekly schedules. This public repository is a **sanitized representative reconstruction**: it preserves the operating model, constraints, assignment logic, workbook structure, and acceptance checks, but it does not contain the real roster, corporate labor targets, historical availability records, or an original posted schedule.

## Program brief

| Field | Detail |
|---|---|
| Business challenge | Produce a feasible weekly schedule across nine restaurant positions while protecting coverage, labor discipline, training supervision, qualifications, and employee constraints. |
| My role | Restaurant manager, operating-rule owner, scheduling decision-maker, and final acceptance reviewer. |
| Users | Store leadership and managers receiving the weekly schedule, coverage checks, hours summary, and decision notes. |
| Scale | One store, 43 representative crew records, seven days, nine positions, and an 11-tab output workbook. |
| Delivery model | Manager-defined operating rules and acceptance decisions; AI-assisted Python implementation and iterative debugging. |
| Status | The historical tooling supported posted schedules. The public code and workbook are a representative, anonymized regeneration for portfolio review. |

## Business problem

The schedule had to do more than fill hours. It needed to respect:

- Opening and closing recovery
- Position qualifications and station capacity
- Minimum shift length and one continuous shift per day
- Performance-priority direction
- Training gates and Certified Trainer shadowing
- Position-rotation expectations
- Register and board concurrency caps
- Labor demand without padding shifts simply to hit hour targets

The operating requirement was a schedule leaders could inspect, explain, adjust, and post - not an opaque optimization result.

## System components

### `schedule_data.py`

The representative data module contains:

- A 43-person pseudonymous roster
- Seven-day availability windows
- Solo-capable positions for each person
- Weekly hour targets
- Closer, no-solo-close, exempt-anchor, performer, fill-in, trainee, and RI flags
- Per-day, per-position demand windows across kitchen, drive-thru, and front counter

### `gen7_engine.py`

The scheduling engine:

1. Splits demand windows into position slots capped at five hours.
2. Prioritizes close slots, then peak slots, then longer blocks.
3. Assigns performers before fill-ins where legal and operationally appropriate.
4. Runs escalating gap-fill passes.
5. Rebalances whole shifts from over-target to under-target qualified employees.
6. Validates training shadows and concurrency limits.
7. Writes assignment, hour, and training JSON files for the workbook builder.

Generated intermediates:

- `assignF.json`
- `hoursF.json`
- `training.json`

### `build8_workbook_builder.py`

The builder reads the engine output and creates an 11-tab workbook with `openpyxl`:

1. Schedule
2. Wednesday position coverage
3. Thursday position coverage
4. Friday position coverage
5. Saturday position coverage
6. Sunday position coverage
7. Monday position coverage
8. Tuesday position coverage
9. Coverage Check
10. Hours Summary and representative labor budget
11. Notes and Decisions

## Encoded operating rules

- **8:00a floor:** no crew shift starts before 8:00a.
- **Three-hour minimum:** assigned shifts must be at least three hours.
- **One continuous shift:** no split shifts for the same person on the same day.
- **Five-hour station cap:** no one holds a single station for more than five hours unless explicitly marked as an exempt anchor.
- **Close-then-open protection:** a 1:00a closer cannot open before 11:00a the following day.
- **Concurrency caps:** no more than two registers and three boards at the same moment.
- **Close capability:** closing assignments respect person-level closer and no-solo-close flags.
- **Training supervision:** every training block requires a Certified Trainer on the same position at the same time.
- **Trainees are additive:** trainees do not count toward required production coverage.
- **Weekday training:** training blocks are not assigned on Saturday or Sunday.
- **Performance priority:** designated performers carry the core schedule; fill-ins cover legal gaps performers cannot take.
- **Target overage cap:** performers may be scheduled up to 12 hours above their representative weekly target.
- **Cross-training gates:** the operating model tracks prerequisite hours before training in a new zone and supervised hours before solo qualification.
- **Restaurant Imaging block:** weekday RI work is separated from production coverage and marked where CIP work occurs.

## Assignment strategy

The public engine is a greedy heuristic with local rebalancing:

1. Build legal candidate lists for each slot.
2. Assign high-priority and hard-to-fill slots first.
3. Prefer qualified performers within hour and availability constraints.
4. Run a relaxed-hours fill for remaining gaps.
5. Run an adjacency-preferring gap closer.
6. Use a last-resort force fill with a 1.5-hour minimum block where required.
7. Run up to 400 rebalancing iterations, moving whole legal shifts to reduce extreme target variance.

It aims for a fast, explainable feasible schedule. It does not prove mathematical optimality.

## Run the representative build

```bash
pip install openpyxl
python gen7_engine.py
python build8_workbook_builder.py
```

Run both files from the repository root. The engine prints diagnostics and writes the three gitignored JSON intermediates. The builder writes:

```text
sample_output/Crew_Schedule_SAMPLE.xlsx
```

## Sample validation result

On the committed representative dataset:

- Uncovered demand: `0.00` hours
- Total scheduled hours: `803.0`
- Every training block has a Certified Trainer on the same station
- Register and board concurrency checks pass
- The workbook includes per-person target comparisons and leadership decision notes

These are validation results for the anonymized sample dataset, not claims about a specific employer week or labor outcome.

## Evidence and provenance

| Item | Evidence status |
|---|---|
| Python engine and builder | Sanitized representative reconstruction of operational tooling. |
| Roster and availability | Pseudonymous and representative; real identities withheld. |
| Sample workbook | Fresh regeneration from public representative data; not an original posted schedule. |
| Coverage and training results | Reproducible on the committed representative dataset. |
| Historical operational use | Leadership-account claim: original versions supported posted schedules. |
| Labor targets | Representative placeholders; corporate forecasts withheld. |

## Repository contents

| Path | Purpose |
|---|---|
| `README.md` | Business context, operating model, evidence status, and limitations. |
| `schedule_data.py` | Public representative roster, availability, demand, and rule flags. |
| `gen7_engine.py` | Assignment, gap-fill, rebalancing, and diagnostics. |
| `build8_workbook_builder.py` | Leadership-facing Excel workbook generation. |
| `sample_output/README.md` | Explains how to regenerate the representative 11-tab workbook locally. |

## Role, contributors, and authorship

Sama Mushtaq supplied the operating context, roster logic, station rules, priorities, qualifications, demand interpretation, manager decision log, and final acceptance decisions. He reviewed diagnostics and workbook outputs before the historical schedules were posted.

The public repository does not claim that every historical implementation line was manually typed by Sama. The leadership artifact is the translation of real operating constraints into requirements, testable rules, exceptions, and accepted outputs.

## AI assistance

Claude was used as a coding partner during the historical build. AI assisted with implementation, iteration, and workbook construction. Sama remained responsible for:

- Operating-rule definition
- Roster and availability interpretation
- Priority and exception decisions
- Training and qualification constraints
- Output review and acceptance
- The decision to post or revise a schedule

AI did not independently determine store policy, employee qualifications, labor targets, or the final posted schedule.

## Confidentiality

- Every public crew name is a pseudonym.
- No real manager roster or manager availability is included.
- Corporate labor targets and forecasts are representative placeholders.
- No third-party scheduling screenshot, source employee record, original posted schedule, or employer document is included.
- The downloadable historical workbook is not committed; reviewers can regenerate a fresh sample from the representative data and builder.
- Representative constraints should be reviewed periodically for triangulation risk even when names are changed.

## Limitations

- Greedy heuristic with local rebalancing; not ILP or CP-SAT and no proof of optimality.
- Single week and single-store operating model.
- Roster and availability are hard-coded rather than imported through a user interface.
- Demand windows and workbook layout are tuned to one restaurant environment.
- Diagnostics are printed acceptance checks, not a full automated test suite.
- Historical source wiring was refactored for publication, so the public code is representative rather than byte-for-byte identical.
- A feasible schedule still requires manager review for real-world information not captured in the model.

## Related

- [Restaurant Imaging corrective-action program](https://github.com/Samamak1/restaurant-imaging-program)
- [Sama Mushtaq program leadership portfolio](https://samamak1.github.io/)

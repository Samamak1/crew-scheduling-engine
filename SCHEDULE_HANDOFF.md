# Crew Schedule, Handoff / How-To-Build

> **Publication note.** This is the working handoff document used to rebuild the schedule in a fresh AI session each week, kept here as an artifact of the process. For publication: all crew names are pseudonyms, manager names other than Sama are generic labels, and corporate labor numbers are redacted. One code change since this was written: the published repo moves the shared data into `schedule_data.py` and both scripts import it, so the `exec(open('gen7.py')...)` wiring and the rename-to-gen7.py step described below no longer apply (see README).

Give this file plus `gen7_engine.py` and `build8_workbook_builder.py` to a new chat and say "continue building my crew schedule from these." Everything needed to rebuild is here.

---

## What these files are
- **gen7_engine.py** = the scheduling engine. It holds the full crew roster, each person's 7-day availability, the position demand per day, all the rules, and the logic that assigns shifts. Running it writes three JSON files (assignF.json, hoursF.json, training.json) and prints diagnostics (gaps, training check, hours, register/board caps).
- **build8_workbook_builder.py** = reads the engine's JSON output and builds the 11-tab Excel workbook.

## How to run (in the new chat's workspace)
1. Put both .py files in the working directory.
2. `python3 gen7_engine.py`  (rename to gen7.py if you like; build8 reads it by the name in its first line, so keep names matched).
3. `python3 build8_workbook_builder.py`  -> writes the workbook to /mnt/user-data/outputs/.
4. Recalc to confirm no formula errors, then present the file.

Note: build8 reads the engine's data section via the line `exec(open('gen7.py')...split('# ==== DATA_END'))`. If you keep the engine named `gen7_engine.py`, update that one line in build8 to match, or just rename the engine to `gen7.py`.

---

## The restaurant
A store in a national quick-service restaurant chain. Profit-share, turnaround, so labor matters (do not pad shifts past demand just to hit targets). HotSchedules store. Schedule is built Wed through Tue. Guests roughly 10a to 12a; crew on at 8:00a at the earliest (no shift starts before 8a).

## Availability source
Google Drive folder "Hotschedules Availability" holds the weekly screenshots (3 images: A-K, K-Sama, S-Z). To read them: download_file_content returns JSON whose inner `content` field is base64 PNG; decode and view. The current week's availability is already transcribed into the CREW dict in the engine, so you only re-read screenshots when a NEW week's availability is posted.

---

## The rules (all already coded in the engine)
- **8:00a floor:** no shift starts before 8a (everyone in at 8a earliest).
- **5-hour position cap:** nobody works one position more than 5 hrs; longer shifts rotate stations. Exempt anchors (may hold a station longer): Avery, Quinn, Tessa, Emil, Felix.
- **3-hour minimum shift; one continuous shift per person per day** (no splits).
- **No close-then-open:** if someone closes (ends 1a), they don't open before 11a next day.
- **Registers capped at 2 at once; boards capped at 3 at once.**
- **No training on Saturday or Sunday.**
- **Every training shift must have a Certified Trainer (CT) working the SAME position at the SAME time.** The trainee is an extra body, not coverage.
- **Cross-training gate:** a person can't start training a new zone until they log 80 hrs in their current zone. A trainee needs 25 training hrs in a zone before working it solo.

## Performance priority (current directive)
Schedule the PERFORMERS first and load them to cover the week; use everyone else only as fill-ins where performers can't cover. Performer overage is capped at target+12 so nobody gets an unreasonable week.
- **PERFORMERS (priority):** Quinn, Avery, Emil, Tessa, Dara, Theo, Harper Nunes, Hugo, Kwame, Rosa, Paloma, Rowan, Ivy, Ezra, Bianca, Beckett, Peyton Calloway, Amina, Marco Renner, Felix, Ansel, Juniper Vale, Genevieve Stroud, Marlowe, Kofi, Silas, Miles (Miles on RI).
- **FILL-INS (only if needed):** Priya, Lena, Wren, Mateo, Selena, Dante, Lionel, Idris, Ja'Nae, Celeste, Bishal, Odette, Simone, Freya.

## Restaurant Imaging (RI)
Miles leads RI, 8:00a to 12:00p, on Wed/Thu/Fri/Mon/Tue (off Sat/Sun, MOD runs a short daily hold those days). CIP (clean ketchup-dispenser line) is folded into Tue and Thu mornings, marked +CIP. RI is out of production coverage. The day-by-day rotation is built off the June standards audit and lives in the separate "RI Program" document. Miles is a performer but his RI block is handled apart from production.

## Training this week
- Camila Sutton: kitchen trainee, ~14 of 25 hrs, shadows a kitchen CT, weekdays only.
- Astrid Larkin: drive-thru trainee, 18 of 25 hrs, shadows a DT CT, weekdays only.
- Celeste Byrne: 3 hrs DT training (shadow a DT CT); can't run a DT position solo until 25 DT hrs; otherwise front counter. Celeste is a fill-in for production.
- Certified Trainers (can supervise): Marco, Miles, Rosa, Avery, Quinn, Tessa, Lionel, Mateo, Ivy, Idris, Theo, Peyton Calloway, Wren.

## Position notes / individual constraints (already in the engine)
- Tessa: front counter / drive-thru only (no kitchen). Paloma: drive-thru only. Silas and Kofi: front counter solo (logging FC hrs toward 80). Celeste: front counter (plus the 3-hr DT training). Peyton Calloway: PM-only.
- Ezra: evenings 5:00p to close every day, open availability Sunday. (HotSchedules shows the opposite, 12a-5p; worth fixing in HS so it stops auto-conflicting.)
- Paloma: on vacation 7/7 to 7/18, so off Tue 7/7.
- A second register runs at lunch on Wed-Sun (helps speed of service and gives front-counter performers hours). Can be dropped to one register to trim labor.

## Roster fill-in zone assumptions (confirm with Sama)
- Simone Deveraux: drive-thru (her certified zone), available Mon and Tue only.
- Freya Adjei and Odette Marchetti: assumed front counter (not in the Crew Journey Tracker). Confirm their real trained zones.

## Not scheduled
Grady Holt, Lachlan Frost, Maren Kovac, Desmond Ruiz, Callum Voss, Estelle Fontaine, Beau Hendrix (departed / vacation / no positions). Managers (Manager A, Sama, Manager B, Manager C, Manager D) run the MOD line, not the crew grid.

---

## Output: the 11-tab workbook
Schedule (grouped Performers / Fill-ins / Trainees / RI), seven day tabs (position coverage timelines), Coverage Check, Hours Summary, Notes & Decisions.

## Formatting preferences
- Documents: hyphens and commas only, no em or en dashes or special characters.
- Resume/application work (separate from scheduling): hyphens and commas only, no tables in the body, operator/builder language.
- Keep replies concise, minimal formatting.

## Current state (week of 7/1 to 7/7)
Built and delivered. Performance-priority applied: performers carry ~667 hrs, fill-ins ~104 hrs, zero coverage gaps, all training CT-covered, registers <=2 and boards <=3. Odette, Simone, Freya sit at 0 hrs this week (lowest-priority fill-ins, not needed). Open dials: set individual performer weekly maxes if some can't take the higher hours, or raise the cap to push fill-ins even lower.

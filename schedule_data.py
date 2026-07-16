"""Shared data module for the crew scheduler.

Holds the representative crew roster, seven-day availability, per-position
demand windows, and rule flags. Both gen7_engine.py (the scheduler) and
build8_workbook_builder.py (the Excel builder) import this module. Publication
consolidates the sample inputs here so the operating rules are reviewable
without preserving historical session wiring or source-system mechanics.

PUBLICATION NOTE: all crew names are pseudonyms and the labor-target numbers
in the workbook builder are representative placeholders. Availability shapes,
position skills, and rule flags are unchanged, so the engine behaves the same.
"""
import math
from collections import defaultdict

DAYS=["Wed","Thu","Fri","Sat","Sun","Mon","Tue"]
WEEKDAYS={"Wed","Thu","Fri","Mon","Tue"}   # no training Sat/Sun
POS_ORDER=["Grill","Board1","Board2","Board3","Toast","DTExp","DTOrd","DTOI","Cash"]
KIT={"Grill","Board1","Board2","Board3","Toast"}; DT={"DTExp","DTOrd","DTOI"}; FC={"Cash"}
ALL=KIT|DT|FC

def fmt(t):
    t2=t if t<24 else t-24
    h=int(t2); m=int(round((t2-h)*60)); ap='a' if h<12 else 'p'; h12=h%12 or 12
    return f"{h12}:{m:02d}{ap}"

# ---- DEMAND (coverage windows) : Cash capped to <=2 concurrent ----
DEMAND={
 "Wed":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(17,20)],"Board3":[],"Toast":[(11,24)],
        "DTExp":[(9,25)],"DTOrd":[(11,22)],"DTOI":[(17,20)],"Cash":[(10,25),(11,15),(17,20)]},
 "Thu":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(11,13),(16,20)],"Board3":[],"Toast":[(11,25)],
        "DTExp":[(8,25)],"DTOrd":[(11,21)],"DTOI":[(17,20)],"Cash":[(10,25),(11,15),(17,20)]},
 "Fri":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(12,14),(16,22)],"Board3":[(17,20)],"Toast":[(11,25)],
        "DTExp":[(8,25)],"DTOrd":[(11,23)],"DTOI":[(12,15),(17,20)],"Cash":[(10,25),(11,15),(17,20)]},
 "Sat":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(11,21)],"Board3":[(12,14),(17,20)],"Toast":[(11,25)],
        "DTExp":[(8,25)],"DTOrd":[(11,23)],"DTOI":[(12,15),(17,20)],"Cash":[(10,25),(11,15),(17,20)]},
 "Sun":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(11,20)],"Board3":[(12,16)],"Toast":[(10.5,25)],
        "DTExp":[(8,25)],"DTOrd":[(11,22)],"DTOI":[(11,15),(16,20)],"Cash":[(10,25),(11,15),(17,20)]},
 "Mon":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[],"Board3":[],"Toast":[(11,24)],
        "DTExp":[(9,25)],"DTOrd":[(11,21)],"DTOI":[],"Cash":[(10,25),(17,20)]},
 "Tue":{"Grill":[(8,25)],"Board1":[(8,25)],"Board2":[(17,20)],"Board3":[],"Toast":[(11,24)],
        "DTExp":[(9,25)],"DTOrd":[(11,21)],"DTOI":[],"Cash":[(10,25),(17,20)]},
}
def snap(x): return round(x*2)/2
def split5(s,e):
    L=e-s
    if L<=5.0: return [(s,e)]
    n=math.ceil(L/5.0); pts=[s]+[snap(s+(L/n)*i) for i in range(1,n)]+[e]
    return [(pts[i],pts[i+1]) for i in range(n) if pts[i+1]>pts[i]]
def is_peak(s,e): return (s<14 and e>11) or (s<20 and e>17)
SLOTS=defaultdict(list)
for d in DAYS:
    for p in POS_ORDER:
        for (s,e) in DEMAND[d].get(p,[]):
            for (a,b) in split5(s,e):
                SLOTS[d].append({"pos":p,"s":a,"e":b,"close":b>=24,"peak":is_peak(a,b),"who":None})

# ---- CREW ----  avail:7-day (8am floor already applied) ; pos=solo-capable ; t=target
# NOTE: every crew name in this file is a pseudonym, the real roster is private (see README).
F=(8,25)
def C(av,pos,t,close,flags=()): return {"avail":av,"pos":set(pos),"target":t,"close":close,"flags":set(flags)}
OFF=None
CREW={
 "Avery":C({"Wed":(16,24),"Thu":(16,24),"Fri":(8,19),"Sat":F,"Sun":F,"Mon":(8,16),"Tue":(8,16)},ALL,35,True,{"exc5"}),
 "Quinn":C({"Wed":F,"Thu":F,"Fri":OFF,"Sat":F,"Sun":(16,24),"Mon":F,"Tue":F},DT|FC,38,True,{"exc5"}),
 "Tessa":C({"Wed":(8,16),"Thu":(8,16),"Fri":(8,16),"Sat":OFF,"Sun":OFF,"Mon":(8,16),"Tue":(8,14)},DT|FC,32,False,{"exc5"}),
 "Emil":C({"Wed":(17,25),"Thu":OFF,"Fri":OFF,"Sat":F,"Sun":OFF,"Mon":(17,25),"Tue":(17,25)},KIT,30,True,{"exc5"}),
 "Priya":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":OFF,"Mon":F,"Tue":F},KIT,26,True),
 "Lena":C({"Wed":(8,23),"Thu":(8,23),"Fri":(8,17),"Sat":(8,23),"Sun":(8,23),"Mon":(8,23),"Tue":(8,23)},DT|FC,30,True),
 "Wren":C({"Wed":(8,21),"Thu":(8,21),"Fri":F,"Sat":F,"Sun":OFF,"Mon":OFF,"Tue":(8,21)},DT|FC,25,True),
 "Theo":C({"Wed":OFF,"Thu":F,"Fri":(8,16),"Sat":F,"Sun":(13,22),"Mon":F,"Tue":F},ALL,24,True),
 "Bianca":C({"Wed":(10,16),"Thu":(10,16),"Fri":(10,16),"Sat":OFF,"Sun":OFF,"Mon":(13,16.5),"Tue":(10,16)},DT|FC,22,False),
 "Dara":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":(16,24),"Mon":F,"Tue":F},DT,25,True),
 "Felix":C({"Wed":OFF,"Thu":OFF,"Fri":OFF,"Sat":(11,22),"Sun":(10,17.5),"Mon":OFF,"Tue":OFF},KIT|DT,16,False,{"exc5"}),
 "Harper":C({"Wed":F,"Thu":F,"Fri":OFF,"Sat":OFF,"Sun":F,"Mon":OFF,"Tue":F},KIT|DT,22,True),
 "Ivy":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":F,"Tue":F},ALL,18,True),
 "Paloma":C({"Wed":(8,23),"Thu":(8,23),"Fri":(8,16),"Sat":(8,16),"Sun":OFF,"Mon":(8,23),"Tue":OFF},DT,20,True),
 "Kofi":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":F,"Tue":F},FC,16,False),
 "Silas":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":(8,17),"Tue":F},FC,14,False,{"no_solo_close"}),
 "Hugo":C({"Wed":(8,16),"Thu":F,"Fri":F,"Sat":F,"Sun":(14,24),"Mon":F,"Tue":F},KIT,24,True,{"new"}),
 "Kwame":C({"Wed":(16,25),"Thu":(16,25),"Fri":(16,25),"Sat":OFF,"Sun":F,"Mon":(16,25),"Tue":(16,25)},KIT,20,True,{"new"}),
 "Mateo":C({"Wed":(16,22.75),"Thu":(16,22.75),"Fri":OFF,"Sat":(8,16),"Sun":(8,16),"Mon":(16,22.75),"Tue":(16,22.75)},KIT|DT,14,False),
 "Peyton":C({"Wed":(16,22),"Thu":(16,20),"Fri":(11,16),"Sat":(11.5,16),"Sun":OFF,"Mon":(11,16),"Tue":(16,22)},KIT|DT,15,False),
 "Beckett":C({"Wed":(8,23),"Thu":(8,23),"Fri":(8,22),"Sat":(8,22),"Sun":(8,22),"Mon":(8,23),"Tue":(8,23)},KIT,16,False,{"new"}),
 "Rowan":C({"Wed":F,"Thu":F,"Fri":OFF,"Sat":OFF,"Sun":OFF,"Mon":OFF,"Tue":OFF},KIT,18,True),
 "Selena":C({"Wed":F,"Thu":F,"Fri":F,"Sat":OFF,"Sun":OFF,"Mon":F,"Tue":F},KIT|DT,22,False),
 "Marco":C({"Wed":(8,20),"Thu":(8,20),"Fri":(8,20),"Sat":OFF,"Sun":OFF,"Mon":OFF,"Tue":(8,20)},DT|FC,16,True),
 "Dante":C({"Wed":F,"Thu":F,"Fri":F,"Sat":(8,12),"Sun":F,"Mon":F,"Tue":F},KIT|DT,18,True),
 "Lionel":C({"Wed":(8,17),"Thu":(8,17),"Fri":(8,17),"Sat":(11,16),"Sun":(11,16),"Mon":(8,17),"Tue":(8,17)},ALL,18,False),
 "Ansel":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":F,"Tue":F},KIT,12,True,{"new"}),
 "Amina":C({"Wed":OFF,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":OFF,"Tue":OFF},DT,16,False,{"new"}),
 "Marlowe":C({"Wed":OFF,"Thu":(16,21),"Fri":OFF,"Sat":OFF,"Sun":(10,19),"Mon":OFF,"Tue":(16,21)},DT,12,False,{"new"}),
 "Idris":C({"Wed":(8,19),"Thu":(8,19),"Fri":(8,19),"Sat":(8,19),"Sun":(8,19),"Mon":(8,19),"Tue":(8,19)},ALL,18,True),
 "Rosa":C({"Wed":(17,24),"Thu":(17,24),"Fri":(8,16),"Sat":(8,16),"Sun":(14,18),"Mon":(17,24),"Tue":(17,24)},ALL,24,True),
 "JaNae":C({"Wed":(8,16),"Thu":(8,16),"Fri":OFF,"Sat":OFF,"Sun":(8,17),"Mon":(8,16),"Tue":(8,16)},DT|FC,14,False),
 "Celeste":C({"Wed":(8,15),"Thu":(8,15),"Fri":(8,15),"Sat":(8,15),"Sun":(8,15),"Mon":(8,15),"Tue":(8,15)},FC,14,False),
 "Bishal":C({"Wed":(19,25),"Thu":(19,25),"Fri":(19,25),"Sat":(19,25),"Sun":OFF,"Mon":OFF,"Tue":OFF},KIT,12,True),
 "Juniper":C({"Wed":OFF,"Thu":(17,22),"Fri":F,"Sat":F,"Sun":(15,22),"Mon":OFF,"Tue":F},FC,14,True,{"new"}),
 "Genevieve":C({"Wed":OFF,"Thu":(13,19),"Fri":F,"Sat":F,"Sun":F,"Mon":F,"Tue":F},FC,18,False,{"new"}),
 "Ezra":C({"Wed":(17,25),"Thu":(17,25),"Fri":(17,25),"Sat":(17,25),"Sun":F,"Mon":(17,25),"Tue":(17,25)},KIT,18,True,{"new"}),
 # ---- fill-ins the user asked to place ----
 "Odette":C({"Wed":F,"Thu":F,"Fri":(8,15.75),"Sat":OFF,"Sun":OFF,"Mon":F,"Tue":F},FC,12,False),
 "Simone":C({"Wed":OFF,"Thu":OFF,"Fri":OFF,"Sat":OFF,"Sun":OFF,"Mon":F,"Tue":F},DT,10,False),
 "Freya":C({"Wed":(8,17),"Thu":(8,17),"Fri":(8,17),"Sat":(8,17),"Sun":(8,17),"Mon":(8,17),"Tue":(8,17)},FC,12,False),
 # ---- trainees (no solo positions; placed only as CT-shadow training) ----
 "Camila":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":(8,13.5),"Mon":F,"Tue":F},set(),14.5,False,{"trainee"}),
 "Astrid":C({"Wed":F,"Thu":F,"Fri":F,"Sat":F,"Sun":F,"Mon":F,"Tue":F},set(),18,False,{"trainee"}),
 # ---- RI lead (handled separately; excluded from production fill) ----
 "Miles":C({"Wed":F,"Thu":F,"Fri":F,"Sat":OFF,"Sun":OFF,"Mon":F,"Tue":F},set(),20,False,{"ri"}),
}
CT={"Marco","Miles","Rosa","Avery","Quinn","Tessa","Lionel","Mateo","Ivy","Idris","Theo","Peyton","Wren"}
EXC5={n for n in CREW if "exc5" in CREW[n]["flags"]}
PROD=[n for n in CREW if "ri" not in CREW[n]["flags"] and "trainee" not in CREW[n]["flags"] and CREW[n]["pos"]]
# ---- PERFORMERS: scheduled first; non-performers only fill what performers can't cover ----
PERFORMERS={"Quinn","Avery","Emil","Tessa","Dara","Theo","Harper","Hugo","Kwame","Rosa",
            "Paloma","Rowan","Ivy","Ezra","Bianca","Beckett","Peyton","Amina","Marco","Felix",
            "Ansel","Juniper","Genevieve","Marlowe","Kofi","Silas","Miles"}
PERF=[n for n in PROD if n in PERFORMERS]
NONPERF=[n for n in PROD if n not in PERFORMERS]

"""Crew scheduling engine (generation 7).

Imports the roster / availability / demand data from schedule_data.py,
assigns crew to position-coverage slots under the house rules (see README),
then writes assignF.json, hoursF.json, training.json for the workbook
builder and prints diagnostics: coverage gaps, training checks, hours vs
target, register/board concurrency.

Run from the repo root:  python3 gen7_engine.py
"""
import json
from collections import defaultdict
from schedule_data import *

# ---- (roster + demand data live in schedule_data.py) ----
# ---- production fill ----
hours=defaultdict(float); day_assign=defaultdict(list); busy=defaultdict(list)
posday=defaultdict(float)  # (name,day,pos)->hrs  (5h cap)
def overlaps(a,b,c,d): return a<d and c<b
def prev_close(name,d):
    i=DAYS.index(d)
    if i==0: return False
    pd=DAYS[i-1]
    return any(e>=24 for (dd,p,s,e) in day_assign[name] if dd==pd)
def can(name,d,pos,s,e,close,peak):
    cr=CREW[name]
    if d not in cr["avail"] or cr["avail"][d] is None: return False
    av=cr["avail"][d]
    if not (av[0]<=s and av[1]>=e): return False
    if pos not in cr["pos"]: return False
    if (e-s)<3.0-1e-9 and (e-s)<(av[1]-av[0]): pass  # min handled post
    for (bs,be) in busy[(name,d)]:
        if overlaps(s,e,bs,be): return False
    if close and not cr["close"]: return False
    if close and "no_solo_close" in cr["flags"]: return False
    if name not in EXC5 and posday[(name,d,pos)]+(e-s)>5.0+1e-9: return False
    if s<11 and prev_close(name,d): return False           # no close-then-open
    cap_h=min(40.0,cr["target"]+12.0) if name in PERFORMERS else cr["target"]+3.0
    if hours[name]+(e-s)>cap_h: return False
    return True
def score(name,d,pos,s,e,close,peak):
    cr=CREW[name]; f=cr["flags"]; sc=(cr["target"]-hours[name])*1.0
    if name in PERFORMERS: sc+=1000   # performers scheduled first
    if close or d=="Sun":
        if name in EXC5: sc+=6
        if name in CT: sc+=3
        if "new" in f: sc-=4
    if peak:
        if "new" in f: sc-=2.5
        if name in CT or name in EXC5: sc+=2
    if pos=="Cash":
        if cr["pos"]=={"Cash"}: sc+=5     # front-counter-only: reserve registers for them
        elif len(cr["pos"])>1: sc-=2.5    # multi-skill: steer toward DT/kitchen instead
        av_end=cr["avail"][d][1]
        if e<=17.5:                       # daytime register slot
            if av_end<=18: sc+=5          # person can only work daytime -> give them daytime register
            elif av_end>=22: sc-=3        # person could close -> save them for evening register
    # keep continuity: prefer same person already adjacent same pos
    for (dd,pp,bs,be) in day_assign[name]:
        if dd==d and pp==pos and (abs(be-s)<1e-6 or abs(e-bs)<1e-6): sc+=3
    ndays=sum(1 for dd in DAYS if cr["avail"].get(dd) is not None)
    if ndays<=2: sc+=6     # scarce availability -> use them on the days they're here
    return sc
def assign_slot(d,sl,n):
    s,e,pos=sl["s"],sl["e"],sl["pos"]
    sl["who"]=n; hours[n]+=(e-s); busy[(n,d)].append((s,e))
    posday[(n,d,pos)]+=(e-s); day_assign[n].append((d,pos,s,e))
# global ordering: close slots (all days) first, then peak, then longest; day index = minor tiebreak
ALLSLOTS=[(d,sl) for d in DAYS for sl in SLOTS[d]]
ALLSLOTS.sort(key=lambda ds:(0 if ds[1]["close"] else 1,0 if ds[1]["peak"] else 1,
                             -(ds[1]["e"]-ds[1]["s"]),DAYS.index(ds[0])))
for d,sl in ALLSLOTS:
    cands=[n for n in PROD if can(n,d,sl["pos"],sl["s"],sl["e"],sl["close"],sl["peak"])]
    if not cands: continue
    assign_slot(d,sl,max(cands,key=lambda n:score(n,d,sl["pos"],sl["s"],sl["e"],sl["close"],sl["peak"])))
# gap-fill: relax the hours cap (allow overflow) to cover remaining demand
def can_relaxed(name,d,pos,s,e,close):
    cr=CREW[name]
    if cr["avail"].get(d) is None: return False
    av=cr["avail"][d]
    if not (av[0]<=s and av[1]>=e): return False
    if pos not in cr["pos"]: return False
    for (bs,be) in busy[(name,d)]:
        if overlaps(s,e,bs,be): return False
    if close and (not cr["close"] or "no_solo_close" in cr["flags"]): return False
    if name not in EXC5 and posday[(name,d,pos)]+(e-s)>5.0+1e-9: return False
    if s<11 and prev_close(name,d): return False
    return True
for d,sl in ALLSLOTS:
    if sl["who"]: continue
    cands=[n for n in PROD if can_relaxed(n,d,sl["pos"],sl["s"],sl["e"],sl["close"])]
    if cands:
        assign_slot(d,sl,min(cands,key=lambda n:(0 if n in PERFORMERS else 1, hours[n]-CREW[n]["target"])))
# rebalance: move a whole movable shift from most-over to an under-target person who can take it
def slots_of(n):
    return [(d,sl) for d,sl in ALLSLOTS if sl["who"]==n]
for _ in range(400):
    over=sorted([n for n in PERF if hours[n]>CREW[n]["target"]+0.5],key=lambda n:-(hours[n]-CREW[n]["target"]))
    moved=False
    for o in over:
        for d,sl in slots_of(o):
            s,e,pos=sl["s"],sl["e"],sl["pos"]
            if e-s<3.0: continue
            takers=[n for n in PERF if n!=o and hours[n]<CREW[n]["target"]-0.5
                    and can(n,d,pos,s,e,sl["close"],sl["peak"])]
            if not takers: continue
            t=min(takers,key=lambda n:hours[n]-CREW[n]["target"])
            # remove from o
            hours[o]-=(e-s); busy[(o,d)].remove((s,e)); posday[(o,d,pos)]-=(e-s)
            day_assign[o].remove((d,pos,s,e))
            assign_slot(d,sl,t); moved=True; break
        if moved: break
    if not moved: break

# ---- final gap-closer: walk each uncovered demand segment, fill with available crew ----
# (relaxes 5h pos-cap & hours-cap as last resort; prefers extending an adjacent same-pos shift,
#  then under-target crew; respects availability, no-overlap, close-capability, no close-then-open)
def prod_cover(d,pos):
    segs=sorted((s,e) for n in PROD for (dd,p,s,e) in day_assign[n] if dd==d and p==pos)
    m=[]
    for a,b in segs:
        if m and a<=m[-1][1]+1e-6: m[-1]=(m[-1][0],max(m[-1][1],b))
        else: m.append((a,b))
    return m
def prod_uncov(d,pos):
    have=prod_cover(d,pos); miss=[]
    for (s,e) in DEMAND[d].get(pos,[]):
        cur=s
        for a,b in have:
            if b<=cur: continue
            if a>cur: miss.append((cur,min(a,e)))
            cur=max(cur,b)
            if cur>=e: break
        if cur<e: miss.append((cur,e))
    return [(a,b) for a,b in miss if b-a>0.25]
def next_conflict(name,d,frm):
    nx=99.0
    for (bs,be) in busy[(name,d)]:
        if bs>=frm-1e-6 and bs<nx: nx=bs
    return nx
def free_at(name,d,pos,s):
    cr=CREW[name]
    if cr["avail"].get(d) is None: return False
    av=cr["avail"][d]
    if not (av[0]<=s+1e-6): return False
    if pos not in cr["pos"]: return False
    if any(overlaps(s,s+0.5,bs,be) for (bs,be) in busy[(name,d)]): return False
    return True
for d in DAYS:
    for pos in POS_ORDER:
        for (gs,ge) in prod_uncov(d,pos):
            cur=gs; guard=0; seg_close=(ge>=24)
            while cur<ge-0.25 and guard<20:
                guard+=1
                cands=[]
                for n in PROD:
                    if not free_at(n,d,pos,cur): continue
                    cr=CREW[n]; av=cr["avail"][d]
                    end=min(ge,av[1],next_conflict(n,d,cur),cur+5.0)
                    if end-cur<3.0: end=min(ge,av[1],next_conflict(n,d,cur),cur+3.0)
                    if end-cur<3.0: continue
                    if seg_close and end>=24 and (not cr["close"] or "no_solo_close" in cr["flags"]): continue
                    if cur<11 and prev_close(n,d): continue
                    adj=any(dd==d and pp==pos and abs(be-cur)<1e-6 for (dd,pp,bs,be) in day_assign[n])
                    cands.append((n,end,adj))
                if not cands: break
                n,end,adj=max(cands,key=lambda c:(50 if c[0] in PERFORMERS else 0)+(1 if c[2] else 0)+(CREW[c[0]]["target"]-hours[c[0]])*0.3+(c[1]-cur)*0.05)
                hours[n]+=(end-cur); busy[(n,d)].append((cur,end))
                posday[(n,d,pos)]+=(end-cur); day_assign[n].append((d,pos,cur,end)); cur=end

# ---- last-resort force-fill: residual gaps, allow short (>=1.5h) blocks, any qualified body ----
MIN_FORCE=1.5
for d in DAYS:
    for pos in POS_ORDER:
        for (gs,ge) in prod_uncov(d,pos):
            cur=gs; guard=0; seg_close=(ge>=24)
            while cur<ge-0.25 and guard<30:
                guard+=1
                cands=[]
                for n in PROD:
                    if not free_at(n,d,pos,cur): continue
                    cr=CREW[n]; av=cr["avail"][d]
                    end=min(ge,av[1],next_conflict(n,d,cur),cur+5.0)
                    if end-cur<MIN_FORCE: continue
                    if seg_close and end>=24 and (not cr["close"] or "no_solo_close" in cr["flags"]): continue
                    if cur<11 and prev_close(n,d): continue
                    cands.append((n,end))
                if not cands: break
                n,end=min(cands,key=lambda c:(0 if c[0] in PERFORMERS else 1, hours[c[0]]-CREW[c[0]]["target"], 1 if c[0] in CT else 0))
                hours[n]+=(end-cur); busy[(n,d)].append((cur,end))
                posday[(n,d,pos)]+=(end-cur); day_assign[n].append((d,pos,cur,end)); cur=end

# merge adjacent same-pos blocks per person/day into one shift entry
def merge(lst):
    by=defaultdict(list)
    for (d,p,s,e) in lst: by[(d,p)].append((s,e))
    out=[]
    for (d,p),iv in by.items():
        iv.sort(); cur=list(iv[0])
        for s,e in iv[1:]:
            if s<=cur[1]+1e-6: cur[1]=max(cur[1],e)
            else: out.append((d,p,cur[0],cur[1])); cur=[s,e]
        out.append((d,p,cur[0],cur[1]))
    return out

# ---- TRAINING placement (trainee shadows a CT on same pos+time; weekdays only) ----
TRAIN=[]
def ct_blocks(day,zone):
    """CT production shifts on `day` whose pos in zone, as (name,pos,s,e)."""
    res=[]
    for n in PROD:
        if n not in CT: continue
        for (d,p,s,e) in merge(day_assign[n]):
            if d==day and p in zone and (e-s)>=3.0: res.append((n,p,s,e))
    res.sort(key=lambda x:-(x[3]-x[2]))
    return res
def place_trainee(tr,zone,need,per=5.0):
    got=0.0
    cr=CREW[tr]
    for d in DAYS:
        if got>=need-1.0: break
        if d not in WEEKDAYS: continue
        if cr["avail"].get(d) is None: continue
        av=cr["avail"][d]
        for (ctn,p,s,e) in ct_blocks(d,zone):
            if got>=need-1.0: break
            ws=max(s,av[0]); we=min(e,av[1])
            if we-ws<3.0: continue
            dur=min(per,we-ws,need-got)
            if dur<3.0: dur=3.0
            if ws+dur>we: dur=we-ws
            if dur<3.0: continue
            # avoid double-booking trainee that day
            if any(overlaps(ws,ws+dur,bs,be) for (bs,be) in busy[(tr,d)]): continue
            TRAIN.append([tr,d,p,round(ws,3),round(ws+dur,3)])
            day_assign[tr].append((d,p,ws,ws+dur)); busy[(tr,d)].append((ws,ws+dur))
            hours[tr]+=dur; got+=dur
            break  # one training block per day
    return got
gJ=place_trainee("Camila",KIT,14.5)
gR=place_trainee("Astrid",DT,18.0)
gN=place_trainee("Celeste",DT,3.0,per=3.0)   # Celeste: 3h DT training (rest of her hrs are FC, added below)

# Celeste also works FC the rest of her availability (solo FC ok). Give her FC slots already filled? 
# She was in PROD (pos=FC) so she already picked up FC coverage. Her DT training is extra above.

# ---- MILES RI + CIP ----  8a-12p on his available weekdays (no Sat/Sun)
RI_DAYS=[d for d in DAYS if CREW["Miles"]["avail"].get(d) is not None]
for d in RI_DAYS:
    day_assign["Miles"].append((d,"RI",8.0,12.0)); hours["Miles"]+=4.0
CIP_DAYS=["Thu","Tue"]  # ketchup dispenser line: Tue + Thu AM (within RI block)

# ---- emit ----
final_assign={n:merge(day_assign[n]) for n in day_assign if day_assign[n]}
hoursF={n:round(sum(e-s for (d,p,s,e) in final_assign.get(n,[])),2) for n in final_assign}
json.dump({n:[[d,p,round(s,3),round(e,3)] for (d,p,s,e) in final_assign[n]] for n in final_assign},
          open("assignF.json","w"))
json.dump(hoursF,open("hoursF.json","w"))
json.dump([[t[0],t[1],t[2],round(t[3],3),round(t[4],3)] for t in TRAIN],open("training.json","w"))

# ---- diagnostics ----
def covered(d,pos):
    segs=sorted([(s,e) for n in final_assign for (dd,p,s,e) in final_assign[n]
                 if dd==d and p==pos and (n,d,pos,round(s,3),round(e,3)) not in {(x[0],x[1],x[2],round(x[3],3),round(x[4],3)) for x in TRAIN}])
    m=[]
    for a,b in segs:
        if m and a<=m[-1][1]+1e-6: m[-1]=(m[-1][0],max(m[-1][1],b))
        else: m.append((a,b))
    return m
def uncov(d,pos):
    have=covered(d,pos); miss=[]
    for (s,e) in DEMAND[d].get(pos,[]):
        cur=s
        for a,b in have:
            if b<=cur: continue
            if a>cur: miss.append((cur,min(a,e)))
            cur=max(cur,b)
            if cur>=e: break
        if cur<e: miss.append((cur,e))
    return [(a,b) for a,b in miss if b-a>0.25]
print("=== GAPS ===")
tg=0
for d in DAYS:
    for pos in POS_ORDER:
        u=uncov(d,pos)
        if u: tg+=sum(b-a for a,b in u); print(f"  {d} {pos}: "+", ".join(f"{fmt(a)}-{fmt(b)}" for a,b in u))
print(f"  total uncovered hrs: {tg:.2f}")
print("\n=== TRAINING (each must have a CT on same pos+time) ===")
for tr,d,p,s,e in TRAIN:
    cts=[n for n in PROD if n in CT and any(dd==d and pp==p and ss<=s+1e-6 and ee>=e-1e-6 for (dd,pp,ss,ee) in merge(day_assign[n]))]
    print(f"  {tr:9s} {d} {p:6s} {fmt(s)}-{fmt(e)}  CT on station: {cts if cts else '*** NONE ***'}")
print(f"  Camila K={gJ:.1f}/14.5  Astrid DT={gR:.1f}/18  Celeste DT={gN:.1f}/3")
print("\n=== HOURS ===")
for n in sorted(final_assign,key=lambda x:-hoursF[x]):
    t=CREW[n]["target"]; h=hoursF[n]; fl="" if abs(h-t)<=3 else ("  LOW" if h<t-3 else "  OVER")
    print(f"  {n:10s} {h:5.1f}/{t:>4.1f}{fl}")
print(f"\n  TOTAL: {sum(hoursF.values()):.1f}")
# cash<=2 and boards<=3 check
print("\n=== CASH>2 / BOARD>3 concurrency check (excl. trainees) ===")
TRSET={(t[0],t[1],t[2],round(t[3],3),round(t[4],3)) for t in TRAIN}
def istr(n,d,p,s,e): return (n,d,p,round(s,3),round(e,3)) in TRSET
bad=0
for d in DAYS:
    for h in [x*0.5 for x in range(16,50)]:
        cc=sum(1 for n in final_assign for (dd,p,s,e) in final_assign[n] if dd==d and p=="Cash" and s<=h<e and not istr(n,d,p,s,e))
        bc=sum(1 for n in final_assign for (dd,p,s,e) in final_assign[n] if dd==d and p in ("Board1","Board2","Board3") and s<=h<e and not istr(n,d,p,s,e))
        if cc>2: print(f"  {d} {fmt(h)} Cash={cc}"); bad+=1
        if bc>3: print(f"  {d} {fmt(h)} Boards={bc}"); bad+=1
print("  ok" if bad==0 else f"  {bad} violations")

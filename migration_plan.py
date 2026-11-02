#!/usr/bin/env python3
"""
Migration planning tool.

Takes planning assumptions (activities, phase durations, test sizes,
environment availability) from parameter tables below, builds a schedule
that respects working days and environment capacity constraints, and
produces a PowerPoint deck with different levels of detail:

  Level 1 - Executive summary: one row per activity with key milestones
  Level 2 - One slide per activity: Gantt chart of phases
  Level 3 - Environment utilization + full legible assumptions tables

Run:  python3 migration_plan.py            -> writes migration_plan.pptx
"""

import argparse
import datetime as dt
import heapq
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu

# ----------------------------------------------------------------------------
#  PARAMETER TABLES  (edit these to change the plan)
# ----------------------------------------------------------------------------

PLAN_START = dt.date(2026, 11, 1)   # first working day of the plan
PLAN_END   = dt.date(2027, 12, 31) # last day of the planning horizon
WORKDAYS_PER_WEEK = 5               # only 5 days a week are used

# (i) Migration activities: name -> type. Types carry default durations.
ACTIVITIES = {
    "Simple1":  "simple",
    "Simple2":  "simple",
    "Simple3":  "simple",
    "Simple4":  "simple",
    "Complex1": "complex",
    "Complex2": "complex",
}

# (ii)/(iii)/(iv)/(v) Phase parameters per activity type.
#   development, application_test, application_test_migrated, inflight_order_test,
#   dress_rehearsal, fut: weeks
#   sit_runs, e2e_runs: number of runs
#   sit_size, e2e_size, dress_rehearsal_size: environment footprint (small/medium/large)
PHASE_PARAMS = {
    "simple": dict(
        development=4,          # weeks
        sit_runs=2,
        sit_size="small",
        e2e_runs=1,
        e2e_size="small",
        application_test=1,                 # weeks
        application_test_migrated=1,        # weeks, runs after application test
        inflight_order_test=2,             # weeks, parallel with application test migrated
        dress_rehearsal=2,                 # weeks, requires an environment
        dress_rehearsal_size="small",
        fut=1,                             # weeks after dress rehearsal
    ),
    "complex": dict(
        development=8,
        sit_runs=5,
        sit_size="large",
        e2e_runs=5,
        e2e_size="medium",
        application_test=10,
        application_test_migrated=10,
        inflight_order_test=4,
        dress_rehearsal=3,
        dress_rehearsal_size="large",
        fut=2,
    ),
}

# (iv) Test sizes: how many days one run occupies the environment.
TEST_SIZES = {
    "small":  dict(sit_days=2, e2e_days=2),
    "medium": dict(sit_days=4, e2e_days=4),
    "large":  dict(sit_days=6, e2e_days=6),
}

# (vi) Environments: name -> capacity and availability window(s).
#   capacity: "small" = can only take small SIT/E2E runs,
#             "all"    = can take any size run.
#   windows:  list of (from_date, to_date) inclusive availability periods.
ENVIRONMENTS = {
    "ENV-A": dict(capacity="all",    windows=[("2026-11-02", "2027-12-31")]),
    "ENV-B": dict(capacity="all",    windows=[("2027-03-01", "2027-12-31")]),
    "ENV-C": dict(capacity="small",  windows=[("2026-11-02", "2027-12-31")]),
    "ENV-D": dict(capacity="small",  windows=[("2027-03-01", "2027-12-31")]),
}

# Minimum break (calendar days) between the end of one run and the start of the
# next run of an activity, by run kind. A SIT run is followed by >= 2 days break,
# an E2E run by >= 3 days break (before the next SIT/E2E/DR run of that activity).
RUN_BREAK_DAYS = {"SIT": 2, "E2E": 3, "DR": 0}

# Minimum separation between any two go-lives (calendar days).
GO_LIVE_SEPARATION_DAYS = 14

# Priority for environment conflicts: 1 = highest, 6 = lowest.
# When two activities compete for the same environment slot, the lower number wins.
# Ties are broken deterministically: earliest-ready run first, then activity name.
ACTIVITY_PRIORITY = {
    "Simple1": 1, "Simple2": 1,
    "Simple3": 3, "Simple4": 3,
    "Complex1": 5, "Complex2": 5,
}

def prio_sort_key(name):
    """Sort activities by priority (1 first), ties by name for stable output."""
    return (ACTIVITY_PRIORITY.get(name, 99), name)

# ----------------------------------------------------------------------------
#  SCHEDULER
# ----------------------------------------------------------------------------

def d(s):
    return dt.date.fromisoformat(s)

def is_workday(day):
    return day.weekday() < WORKDAYS_PER_WEEK

def next_workday(day):
    while not is_workday(day):
        day += dt.timedelta(days=1)
    return day

def add_workdays(start, n):
    """n working days after start (start itself counts as day 1 if workday)."""
    day = next_workday(start)
    count = 1
    while count < n:
        day += dt.timedelta(days=1)
        day = next_workday(day)
        count += 1
    return day

@dataclass
class Run:
    activity: str
    kind: str          # "SIT" or "E2E"
    size: str
    start: dt.date
    end: dt.date
    env: str

@dataclass
class ActivityPlan:
    name: str
    type: str
    phases: dict = field(default_factory=dict)   # phase -> (start, end)
    runs: list = field(default_factory=list)
    e2e_end: dt.date = None                       # end of last SIT/E2E run
    @property
    def ready(self):
        return max(self.phases["Application test migrated"][1],
                   self.phases["Inflight order test"][1])
    @property
    def golive(self):
        return self.phases["Go-live"][1]

class Scheduler:
    def __init__(self):
        self.env_bookings = {name: [] for name in ENVIRONMENTS}
        self.runs = []

    def env_ok(self, env, size, day):
        cap = ENVIRONMENTS[env]["capacity"]
        if cap != "all" and size != "small":
            return False
        for frm, to in ENVIRONMENTS[env]["windows"]:
            if d(frm) <= day <= d(to):
                return True
        return False

    def env_free(self, env, days):
        booked = {b for _, b in self.env_bookings[env]}
        return all(day not in booked for day in days)

    def book_run(self, activity, kind, size, earliest, n_days=None):
        """Find earliest slot of consecutive working days in a capable, free env."""
        if n_days is None:
            n_days = TEST_SIZES[size]["sit_days"] if kind == "SIT" else TEST_SIZES[size]["e2e_days"]
        # A run must fall inside an environment availability window, so there is
        # no point searching past the last available day of any window.
        last_available = max(d(to) for env in ENVIRONMENTS.values()
                             for _, to in env["windows"])
        day = next_workday(earliest)
        while day <= last_available:
            days = []
            cur = day
            for _ in range(n_days):
                cur = next_workday(cur)
                days.append(cur)
                cur += dt.timedelta(days=1)
            candidates = [e for e in ENVIRONMENTS
                          if self.env_ok(e, size, day)
                          and self.env_ok(e, size, days[-1])
                          and self.env_free(e, days)]
            if candidates:
                env = sorted(candidates)[0]
                for dayx in days:
                    self.env_bookings[env].append((activity, dayx))
                run = Run(activity, kind, size, days[0], days[-1], env)
                self.runs.append(run)
                return run
            day = next_workday(day + dt.timedelta(days=1))
        # Search deadline reached without a slot: explain why instead of looping forever.
        capable = [e for e in ENVIRONMENTS
                   if ENVIRONMENTS[e]["capacity"] == "all" or size == "small"]
        raise RuntimeError(
            f"Cannot schedule {kind} run ({size}, {n_days} days) for {activity}: "
            f"no capable, available, free environment was found before "
            f"{fmt(last_available)} (the last day of any environment window). "
            f"Capable environments by size: {capable or 'none'}. "
            f"Check ENVIRONMENTS windows/capacity, TEST_SIZES, PLAN_START and the "
            f"number of runs in PHASE_PARAMS.")

    def plan(self):
        """Schedule all activities.

        Runs from all activities compete globally for environments. Runs are
        booked in priority order (1 = highest); within the same priority, the
        run that is ready earliest books first, ties broken by activity name.
        A run's earliest start is the day after its activity's previous phase/run
        ends (development -> SIT runs -> E2E runs are sequential per activity).
        """
        plans = {}
        heap = []   # entries: (prio, ready_date, name, kind, size, run_seq)

        # Phase 1: development phases are unconstrained (no environment needed)
        for name in ACTIVITIES:
            p = PHASE_PARAMS[ACTIVITIES[name]]
            a = ActivityPlan(name, ACTIVITIES[name])
            start = next_workday(PLAN_START)
            end = add_workdays(start, p["development"] * WORKDAYS_PER_WEEK)
            a.phases["Development"] = (start, end)
            plans[name] = a
            # First run (SIT) becomes ready the day after development ends
            if p["sit_runs"]:
                heap.append((ACTIVITY_PRIORITY.get(name, 99),
                             next_workday(end + dt.timedelta(days=1)),
                             name, "SIT", p["sit_size"], 1))
        heapq.heapify(heap)

        # Phase 2: book environment runs in global priority order
        remaining = {name: dict(sit=PHASE_PARAMS[ACTIVITIES[name]]["sit_runs"],
                                e2e=PHASE_PARAMS[ACTIVITIES[name]]["e2e_runs"],
                                sit_size=PHASE_PARAMS[ACTIVITIES[name]]["sit_size"],
                                e2e_size=PHASE_PARAMS[ACTIVITIES[name]]["e2e_size"])
                     for name in ACTIVITIES}
        while heap:
            entry = heapq.heappop(heap)
            prio, ready, name, kind, size, seq = entry[:6]
            n_days = entry[6] if len(entry) > 6 else None
            run = self.book_run(name, kind, size, ready, n_days=n_days)
            plans[name].runs.append(run)
            # Chain the next run of this activity (SIT runs, then E2E runs),
            # honouring the minimum break after this run's kind.
            break_days = RUN_BREAK_DAYS.get(kind, 0)
            nxt_ready = next_workday(run.end + dt.timedelta(days=1 + break_days))
            rem = remaining[name]
            if kind == "SIT":
                rem["sit"] -= 1
                if rem["sit"] > 0:
                    heapq.heappush(heap, (prio, nxt_ready, name, "SIT", rem["sit_size"], seq + 1))
                elif rem["e2e"] > 0:
                    heapq.heappush(heap, (prio, nxt_ready, name, "E2E", rem["e2e_size"], seq + 1))
            elif kind == "E2E":
                rem["e2e"] -= 1
                if rem["e2e"] > 0:
                    heapq.heappush(heap, (prio, nxt_ready, name, "E2E", rem["e2e_size"], seq + 1))
                else:
                    # Last SIT/E2E run done: remember it, compute readiness, and
                    # queue the dress rehearsal (environment run) for phase 2.
                    p = PHASE_PARAMS[ACTIVITIES[name]]
                    a = plans[name]
                    a.e2e_end = run.end
                    app_start = nxt_ready
                    app_end = add_workdays(app_start, p["application_test"] * WORKDAYS_PER_WEEK)
                    a.phases["Application test"] = (app_start, app_end)
                    mig_end = add_workdays(next_workday(app_end + dt.timedelta(days=1)),
                                           p["application_test_migrated"] * WORKDAYS_PER_WEEK)
                    a.phases["Application test migrated"] = (
                        next_workday(app_end + dt.timedelta(days=1)), mig_end)
                    inf_end = add_workdays(next_workday(app_end + dt.timedelta(days=1)),
                                           p["inflight_order_test"] * WORKDAYS_PER_WEEK)
                    a.phases["Inflight order test"] = (
                        next_workday(app_end + dt.timedelta(days=1)), inf_end)
                    ready = max(mig_end, inf_end)
                    dr_weeks = p["dress_rehearsal"]
                    heapq.heappush(heap, (prio, next_workday(ready + dt.timedelta(days=1)),
                                          name, "DR", p["dress_rehearsal_size"], seq + 1,
                                          dr_weeks * WORKDAYS_PER_WEEK))

        # Phase 3: FUT after dress rehearsal, then go-live with separation
        last_golive = None
        for name in sorted(ACTIVITIES, key=prio_sort_key):
            p = PHASE_PARAMS[ACTIVITIES[name]]
            a = plans[name]
            dr_runs = [r for r in a.runs if r.kind == "DR"]
            dr_end = max(r.end for r in dr_runs)
            fut_start = next_workday(dr_end + dt.timedelta(days=1))
            fut_end = add_workdays(fut_start, p["fut"] * WORKDAYS_PER_WEEK)
            a.phases["FUT"] = (fut_start, fut_end)
            # Go-live: earliest the next working day after FUT, but at least
            # GO_LIVE_SEPARATION_DAYS after any previously scheduled go-live.
            earliest = next_workday(fut_end + dt.timedelta(days=1))
            if last_golive is not None:
                earliest = max(earliest,
                               next_workday(last_golive + dt.timedelta(days=GO_LIVE_SEPARATION_DAYS)))
            a.phases["Go-live"] = (earliest, earliest)
            last_golive = earliest
        return plans

# ----------------------------------------------------------------------------
#  POWERPOINT OUTPUT
# ----------------------------------------------------------------------------

BAR_COLORS = {
    "Development":          "4472C4",
    "SIT":                  "ED7D31",
    "E2E":                  "FFC000",
    "Application test":     "70AD47",
    "Application test migrated": "9BD58A",
    "Inflight order test":  "7030A0",
    "Dress rehearsal":      "E060A0",
    "FUT":                  "C00000",
    "Go-live":              "000000",
}
PHASES = list(BAR_COLORS)

def fmt(day):
    return day.strftime("%d-%b-%y")

def add_title(slide, text, sub=None):
    box = slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(9.2), Inches(0.7))
    tf = box.text_frame
    tf.text = text
    tf.paragraphs[0].runs[0].font.size = Pt(24)
    tf.paragraphs[0].runs[0].font.bold = True
    if sub:
        p = tf.add_paragraph()
        r = p.add_run()
        r.text = sub
        r.font.size = Pt(12)

def add_table(slide, rows, cols, left, top, width, height):
    shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top),
                                   Inches(width), Inches(height))
    return shape.table

def set_cell(cell, text, bold=False, size=10):
    cell.text = str(text)
    para = cell.text_frame.paragraphs[0]
    para.font.size = Pt(size)
    para.font.bold = bold

def slide_exec_summary(prs, plans):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_title(slide, "Migration Plan - Executive Summary",
              "Level 1: one line per activity, key milestones")
    ordered = [plans[n] for n in sorted(ACTIVITY_PRIORITY, key=prio_sort_key) if n in plans]
    rows = len(ordered) + 1
    table = add_table(slide, rows, 10, 0.4, 1.1, 9.2, 0.4 * rows)
    headers = ["Activity", "Type", "Dev done", "Tests done", "App test done",
               "Ready", "Dress rehearsal", "FUT done", "Go-live", "Total weeks"]
    for i, h in enumerate(headers):
        set_cell(table.cell(0, i), h, bold=True)
    for r, a in enumerate(ordered, start=1):
        total = (a.golive - PLAN_START).days / 7
        vals = [a.name, a.type,
                fmt(a.phases["Development"][1]),
                fmt(max((x.end for x in a.runs if x.kind in ("SIT", "E2E")), default=PLAN_START)),
                fmt(a.phases["Application test"][1]),
                fmt(a.ready), fmt(max(x.end for x in a.runs if x.kind == "DR")),
                fmt(a.phases["FUT"][1]), fmt(a.golive), f"{total:.0f}"]
        for c, v in enumerate(vals):
            set_cell(table.cell(r, c), v)
    return slide

def slide_activity(prs, a):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_title(slide, f"{a.name} ({a.type}) - Detailed Plan",
              "Level 2: phases and environment bookings")

    chart_top = Inches(1.2)
    chart_left = Inches(2.4)
    chart_width = Inches(6.8)
    plan_days = (a.golive - PLAN_START).days + 3

    def x_of(day):
        return Emu(int(chart_left + chart_width * ((day - PLAN_START).days / plan_days)))

    # Date axis labels
    axis = slide.shapes.add_textbox(Inches(2.4), Inches(1.0), Inches(6.8), Inches(0.25))
    tf = axis.text_frame
    tf.text = fmt(PLAN_START)
    tf.paragraphs[0].runs[0].font.size = Pt(8)
    p = tf.add_paragraph()
    tf2 = slide.shapes.add_textbox(x_of(a.golive) - Inches(0.5), Inches(1.0),
                                   Inches(1.0), Inches(0.25))
    tf2.text_frame.text = fmt(a.golive)
    tf2.text_frame.paragraphs[0].runs[0].font.size = Pt(8)

    # Merge sequential phases and individual runs into timeline bars
    bars = []
    for ph in PHASES:
        if ph in a.phases:
            s, e = a.phases[ph]
            bars.append((ph, s, e, False))
    for run in a.runs:
        s, e = run.start, run.end
        label = f"{run.kind} {run.size} @ {run.env}"
        bars.append((label, s, e, True))
    bars.sort(key=lambda b: (b[1], b[2]))

    row_h = Inches(0.28)
    for i, (label, s, e, is_run) in enumerate(bars):
        y = Emu(int(chart_top + i * row_h))
        # row label
        lb = slide.shapes.add_textbox(Inches(0.4), y, Inches(1.9), Inches(0.25))
        lb.text_frame.text = label
        lb.text_frame.paragraphs[0].runs[0].font.size = Pt(8)
        color = BAR_COLORS.get(label.split(" ")[0] if is_run else label, "808080")
        x1, x2 = x_of(s), x_of(e) + Inches(0.12)
        rect = slide.shapes.add_shape(1, x1, y, Emu(int(x2 - x1)), Inches(0.18))
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor.from_string(color)
        rect.line.fill.background()

    # Key milestones box
    n = len(bars)
    box_top = Emu(int(chart_top + (n + 2) * row_h))
    tb = slide.shapes.add_textbox(Inches(0.4), box_top, Inches(9.2), Inches(1.2))
    tf = tb.text_frame
    tf.text = (f"Ready for migration: {fmt(a.ready)}    "
               f"Go-live preparation complete: {fmt(a.golive)}")
    tf.paragraphs[0].runs[0].font.size = Pt(12)
    tf.paragraphs[0].runs[0].font.bold = True

def slide_env_utilization(prs, sched):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_title(slide, "Environment Utilization",
              "Level 3: days booked per environment")
    from collections import Counter
    envs = sorted(ENVIRONMENTS)
    rows = len(envs) + 1
    table = add_table(slide, rows, 5, 0.4, 1.1, 9.2, 0.35 * rows)
    headers = ["Environment", "Capacity", "Available from", "Available to", "Booked days"]
    for i, h in enumerate(headers):
        set_cell(table.cell(0, i), h, bold=True)
    for r, e in enumerate(envs, start=1):
        booked = Counter(act for act, _ in sched.env_bookings[e])
        total = len(sched.env_bookings[e])
        detail = ", ".join(f"{k}:{v}" for k, v in sorted(booked.items())) or "-"
        set_cell(table.cell(r, 0), e)
        set_cell(table.cell(r, 1), ENVIRONMENTS[e]["capacity"])
        wins = ENVIRONMENTS[e]["windows"]
        set_cell(table.cell(r, 2), fmt(d(wins[0][0])))
        set_cell(table.cell(r, 3), fmt(d(wins[-1][1])))
        set_cell(table.cell(r, 4), f"{total} ({detail})")
    return slide

def slide_assumptions(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_title(slide, "Planning Assumptions",
              "Level 3: all parameter values used by the plan")
    row = 1.1
    # Activities table
    table = add_table(slide, len(ACTIVITIES) + 1, 9, 0.4, row, 9.2,
                      0.22 * (len(ACTIVITIES) + 1))
    headers = ["Activity", "Type", "Dev (wk)", "SIT runs", "SIT size",
               "E2E runs", "E2E size", "App test (wk)", "Inflight (wk)"]
    for i, h in enumerate(headers):
        set_cell(table.cell(0, i), h, bold=True, size=8)
    for r, (name, typ) in enumerate(ACTIVITIES.items(), start=1):
        p = PHASE_PARAMS[typ]
        vals = [name, typ, p["development"], p["sit_runs"], p["sit_size"],
                p["e2e_runs"], p["e2e_size"], p["application_test"],
                p["inflight_order_test"]]
        for c, v in enumerate(vals):
            set_cell(table.cell(r, c), v, size=8)
    row += 0.22 * (len(ACTIVITIES) + 1) + 0.3

    # Test sizes table
    nsz = len(TEST_SIZES)
    table = add_table(slide, nsz + 1, 3, 0.4, row, 4.5, 0.22 * (nsz + 1))
    for i, h in enumerate(["Test size", "SIT days/run", "E2E days/run"]):
        set_cell(table.cell(0, i), h, bold=True, size=8)
    for r, (sz, v) in enumerate(TEST_SIZES.items(), start=1):
        set_cell(table.cell(r, 0), sz, size=8)
        set_cell(table.cell(r, 1), v["sit_days"], size=8)
        set_cell(table.cell(r, 2), v["e2e_days"], size=8)

    # Environments table
    table = add_table(slide, len(ENVIRONMENTS) + 1, 4, 5.1, row, 4.5,
                      0.22 * (len(ENVIRONMENTS) + 1))
    for i, h in enumerate(["Environment", "Capacity", "Available from", "Available to"]):
        set_cell(table.cell(0, i), h, bold=True, size=8)
    for r, (e, v) in enumerate(ENVIRONMENTS.items(), start=1):
        set_cell(table.cell(r, 0), e, size=8)
        set_cell(table.cell(r, 1), v["capacity"], size=8)
        set_cell(table.cell(r, 2), fmt(d(v["windows"][0][0])), size=8)
        set_cell(table.cell(r, 3), fmt(d(v["windows"][-1][1])), size=8)
    row += max(0.22 * (nsz + 1), 0.22 * (len(ENVIRONMENTS) + 1)) + 0.3

    # General assumptions text
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(row), Inches(9.2), Inches(1.5))
    tf = tb.text_frame
    lines = [
        f"Plan starts: {fmt(PLAN_START)}   |   Working days per week: {WORKDAYS_PER_WEEK}",
        "Per activity: Development -> SIT runs -> E2E runs -> Application test -> "
        "Application test migrated / Inflight order test (parallel) -> Dress rehearsal "
        "(environment) -> FUT -> Go-live are sequential.",
        "Inflight order test runs in parallel with Application test migrated.",
        "SIT and E2E runs book a capable, available environment for consecutive working days.",
        "Minimum break between runs of an activity: "
        + ", ".join(f"{k} {v}d" for k, v in RUN_BREAK_DAYS.items() if v) + ".",
        "Environment conflicts resolve by priority (1 = highest, 6 = lowest): "
        + ", ".join(f"{n}={p}" for n, p in sorted(ACTIVITY_PRIORITY.items(), key=prio_sort_key)) + ".",
        "Go-lives are separated by at least "
        f"{GO_LIVE_SEPARATION_DAYS} days and scheduled in priority order.",
    ]
    tf.text = "\n".join(lines)
    for p in tf.paragraphs:
        p.runs[0].font.size = Pt(9) if p.runs else 9

def build_deck(plans, sched, path):
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide_exec_summary(prs, plans)
    for name in sorted(ACTIVITY_PRIORITY, key=prio_sort_key):
        if name in plans:
            slide_activity(prs, plans[name])
    slide_env_utilization(prs, sched)
    slide_assumptions(prs)
    prs.save(path)

# ----------------------------------------------------------------------------
#  TEXT OUTPUT (console)
# ----------------------------------------------------------------------------

def print_overview(plans):
    """Level 1 - one line per activity, key milestones."""
    print("=" * 100)
    print(f"MIGRATION PLAN OVERVIEW   horizon: {fmt(PLAN_START)} -> {fmt(PLAN_END)}   "
          f"({WORKDAYS_PER_WEEK} working days/week)")
    print("=" * 100)
    hdr = f"  {'Activity':<10s} {'Prio':<4s} {'Type':<8s} {'Start':<10s} {'Dev done':<10s} " \
          f"{'SIT done':<10s} {'E2E done':<10s} {'App done':<10s} {'AppMig done':<10s} " \
          f"{'Ready':<10s} {'DR done':<10s} {'FUT done':<10s} {'Go-live':<12s} {'Weeks':>5s}"
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    for name in sorted(ACTIVITY_PRIORITY, key=prio_sort_key):
        if name not in plans:
            continue
        a = plans[name]
        start = min(s for s, _ in a.phases.values())
        def last_end(kind):
            ends = [r.end for r in a.runs if r.kind == kind]
            return fmt(max(ends)) if ends else "-"
        weeks = (a.golive - start).days // 7
        print(f"  {a.name:<10s} {ACTIVITY_PRIORITY.get(name, '?'):<4d} {a.type:<8s} {fmt(start):<10s} "
              f"{fmt(a.phases['Development'][1]):<10s} {last_end('SIT'):<10s} {last_end('E2E'):<10s} "
              f"{fmt(a.phases['Application test'][1]):<10s} "
              f"{fmt(a.phases['Application test migrated'][1]):<10s} "
              f"{fmt(a.ready):<10s} {last_end('DR'):<10s} {fmt(a.phases['FUT'][1]):<10s} "
              f"{fmt(a.golive):<12s} {weeks:>4d}")
    # Horizon check
    late = [a.name for a in plans.values() if a.golive > PLAN_END]
    if late:
        print(f"\n  !! WARNING: activities finishing after the planning horizon: {', '.join(late)}")
    print()

def print_env_use(sched):
    """Environment use over time: booked environment-days per month."""
    print("=" * 100)
    print("ENVIRONMENT USE OVER TIME  (booked environment-days per month; "
          "n/a = outside availability window)")
    print("=" * 100)

    # Month buckets from plan start to horizon
    months = []
    y, m = PLAN_START.year, PLAN_START.month
    while (y, m) <= (PLAN_END.year, PLAN_END.month):
        months.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1

    # Booked days per environment per month
    use = {e: {(yy, mm): 0 for (yy, mm) in months} for e in ENVIRONMENTS}
    for env, bookings in sched.env_bookings.items():
        for _, day in bookings:
            key = (day.year, day.month)
            if key in use[env]:
                use[env][key] += 1

    def avail_in_month(env, ym):
        y, m = ym
        for frm, to in ENVIRONMENTS[env]["windows"]:
            if not (d(to).year < y or (d(to).year == y and d(to).month < m)) and \
               not (d(frm).year > y or (d(frm).year == y and d(frm).month > m)):
                return True
        return False

    hdr = "  {:<10s} {:<7s}".format("Environm.", "Cap") \
        + "".join(f"{dt.date(y, m, 1).strftime('%b%y'):>6s}" for (y, m) in months) \
        + f"{'Total':>7s}"
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    grand = 0
    for e in sorted(ENVIRONMENTS):
        total = len(sched.env_bookings[e])
        grand += total
        row = "  {:<10s} {:<7s}".format(e, ENVIRONMENTS[e]["capacity"])
        for ym in months:
            if avail_in_month(e, ym):
                row += f"{use[e][ym]:>6d}"
            else:
                row += f"{'n/a':>6s}"
        row += f"{total:>7d}"
        print(row)
    print("  " + "-" * (len(hdr) - 2))
    print(f"  {'Total':<18s}".format("") + " " * (6 * len(months)) + f"{grand:>7d}")
    print()

def print_plans(plans, sched):
    print_overview(plans)
    print_env_use(sched)
    print("=" * 100)
    print("DETAIL")
    print("=" * 100)
    for name in sorted(ACTIVITY_PRIORITY, key=prio_sort_key):
        if name not in plans:
            continue
        a = plans[name]
        print(f"\n--- {a.name} ({a.type}, prio {ACTIVITY_PRIORITY.get(name, '?')}) ---")
        for ph in PHASES:
            if ph in a.phases:
                s, e = a.phases[ph]
                print(f"  {ph:24s} {fmt(s)} -> {fmt(e)}")
        for run in a.runs:
            print(f"  {run.kind} {run.size:8s} @ {run.env:6s} {fmt(run.start)} -> {fmt(run.end)}")
        print(f"  {'Ready for migration':24s} {fmt(a.ready)}")
        print(f"  {'Go-live':24s} {fmt(a.golive)}")
    print("\nEnvironment bookings:")
    for e in sorted(sched.env_bookings):
        print(f"  {e}: {len(sched.env_bookings[e])} days booked")

    print("\nAssumptions:")
    print(f"  Activities: {len(ACTIVITIES)} "
          f"({sum(1 for t in ACTIVITIES.values() if t == 'simple')} simple, "
          f"{sum(1 for t in ACTIVITIES.values() if t == 'complex')} complex)")
    for typ, p in PHASE_PARAMS.items():
        print(f"  [{typ}] dev {p['development']}wk | SIT {p['sit_runs']}x{p['sit_size']} "
              f"({TEST_SIZES[p['sit_size']]['sit_days']}d) | E2E {p['e2e_runs']}x{p['e2e_size']} "
              f"({TEST_SIZES[p['e2e_size']]['e2e_days']}d) | app test {p['application_test']}wk "
              f"| app test migrated {p['application_test_migrated']}wk "
              f"| inflight {p['inflight_order_test']}wk | dress rehearsal {p['dress_rehearsal']}wk "
              f"({p['dress_rehearsal_size']}) | FUT {p['fut']}wk")
    for e, v in ENVIRONMENTS.items():
        wins = "; ".join(f"{frm} to {to}" for frm, to in v["windows"])
        print(f"  {e}: capacity={v['capacity']}, available {wins}")

# ----------------------------------------------------------------------------
#  MAIN
# ----------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Migration planning tool")
    ap.add_argument("-o", "--output", default="migration_plan.pptx")
    args = ap.parse_args()

    sched = Scheduler()
    plans = sched.plan()
    print_plans(plans, sched)
    #build_deck(plans, sched, args.output)
    #print(f"\nPowerPoint deck written to: {args.output}")

if __name__ == "__main__":
    main()

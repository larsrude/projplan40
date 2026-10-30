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
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu

# ----------------------------------------------------------------------------
#  PARAMETER TABLES  (edit these to change the plan)
# ----------------------------------------------------------------------------

PLAN_START = dt.date(2025, 1, 6)   # first working day of the plan
WORKDAYS_PER_WEEK = 5              # only 5 days a week are used

# (i) Migration activities: name -> type. Types carry default durations.
ACTIVITIES = {
    "Simple1":  "simple",
    "Simple2":  "simple",
    "Simple3":  "simple",
    "Simple4":  "simple",
    "Complex1": "complex",
    "Complex2": "complex",
}

# (ii)/(iii) Phase parameters per activity type.
#   development, application_test, inflight_order_test, go_live_prep: weeks
#   sit_runs, e2e_runs: number of runs
#   sit_size, e2e_size: environment footprint (small/medium/large/xlarge)
PHASE_PARAMS = {
    "simple": dict(
        development=4,          # weeks
        sit_runs=2,
        sit_size="small",
        e2e_runs=1,
        e2e_size="small",
        application_test=3,     # weeks
        inflight_order_test=2,  # weeks, runs in parallel with application test
        go_live_prep=2,         # weeks after readiness
    ),
    "complex": dict(
        development=8,
        sit_runs=3,
        sit_size="large",
        e2e_runs=2,
        e2e_size="medium",
        application_test=5,
        inflight_order_test=4,
        go_live_prep=3,
    ),
}

# (iv) Test sizes: how many days one run occupies the environment.
TEST_SIZES = {
    "small":  dict(sit_days=5,  e2e_days=5),
    "medium": dict(sit_days=10, e2e_days=10),
    "large":  dict(sit_days=15, e2e_days=15),
    "xlarge": dict(sit_days=25, e2e_days=20),
}

# (vi) Environments: name -> capacity and availability window(s).
#   capacity: "small" = can only take small SIT/E2E runs,
#             "all"    = can take any size run.
#   windows:  list of (from_date, to_date) inclusive availability periods.
ENVIRONMENTS = {
    "ENV-A": dict(capacity="all",    windows=[("2025-01-06", "2025-12-31")]),
    "ENV-B": dict(capacity="all",    windows=[("2025-03-03", "2025-12-31")]),
    "ENV-C": dict(capacity="small",  windows=[("2025-01-06", "2025-06-30")]),
    "ENV-D": dict(capacity="small",  windows=[("2025-02-03", "2025-09-30")]),
}

# Priority order in which activities compete for environments.
ACTIVITY_PRIORITY = ["Complex1", "Complex2", "Simple1", "Simple2", "Simple3", "Simple4"]

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
    @property
    def ready(self):
        return max(self.phases["Application test"][1],
                   self.phases["Inflight order test"][1])
    @property
    def golive(self):
        return self.phases["Go-live preparation"][1]

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

    def book_run(self, activity, kind, size, earliest):
        """Find earliest slot of consecutive working days in a capable, free env."""
        n_days = TEST_SIZES[size][kind.lower().replace("e2e", "e2e") + "_days"] \
            if kind == "SIT" else TEST_SIZES[size]["e2e_days"]
        day = next_workday(earliest)
        while True:
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

    def plan(self):
        plans = {}
        for name in ACTIVITY_PRIORITY:
            if name not in ACTIVITIES:
                continue
            p = PHASE_PARAMS[ACTIVITIES[name]]
            a = ActivityPlan(name, ACTIVITIES[name])
            cursor = PLAN_START

            # Development (weeks)
            end = add_workdays(cursor, p["development"] * WORKDAYS_PER_WEEK)
            a.phases["Development"] = (cursor, end)
            cursor = end + dt.timedelta(days=1)

            # SIT runs (sequential, environment constrained)
            for _ in range(p["sit_runs"]):
                run = self.book_run(name, "SIT", p["sit_size"], cursor)
                a.runs.append(run)
                cursor = run.end + dt.timedelta(days=1)

            # End-to-end test runs (sequential, environment constrained)
            for _ in range(p["e2e_runs"]):
                run = self.book_run(name, "E2E", p["e2e_size"], cursor)
                a.runs.append(run)
                cursor = run.end + dt.timedelta(days=1)

            # Application test and inflight order test in parallel
            app_end = add_workdays(cursor, p["application_test"] * WORKDAYS_PER_WEEK)
            a.phases["Application test"] = (cursor, app_end)
            inf_end = add_workdays(cursor, p["inflight_order_test"] * WORKDAYS_PER_WEEK)
            a.phases["Inflight order test"] = (cursor, inf_end)

            # Go-live preparation after readiness
            ready = max(app_end, inf_end)
            cursor = next_workday(ready + dt.timedelta(days=1))
            gl_end = add_workdays(cursor, p["go_live_prep"] * WORKDAYS_PER_WEEK)
            a.phases["Go-live preparation"] = (cursor, gl_end)

            plans[name] = a
        return plans

# ----------------------------------------------------------------------------
#  POWERPOINT OUTPUT
# ----------------------------------------------------------------------------

BAR_COLORS = {
    "Development":          "4472C4",
    "SIT":                  "ED7D31",
    "E2E":                  "FFC000",
    "Application test":     "70AD47",
    "Inflight order test":  "7030A0",
    "Go-live preparation":  "C00000",
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
    ordered = [plans[n] for n in ACTIVITY_PRIORITY if n in plans]
    rows = len(ordered) + 1
    table = add_table(slide, rows, 7, 0.4, 1.1, 9.2, 0.4 * rows)
    headers = ["Activity", "Type", "Dev done", "Tests done", "Ready for migration",
               "Go-live prep done", "Total weeks"]
    for i, h in enumerate(headers):
        set_cell(table.cell(0, i), h, bold=True)
    for r, a in enumerate(ordered, start=1):
        total = (a.golive - PLAN_START).days / 7
        vals = [a.name, a.type,
                fmt(a.phases["Development"][1]),
                fmt(max(a.runs[-1].end for a in [a] if a.runs) if a.runs else PLAN_START),
                fmt(a.ready), fmt(a.golive), f"{total:.0f}"]
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
        "Go-live preparation are sequential.",
        "Inflight order test runs in parallel with Application test (starts same day).",
        "SIT and E2E runs book a capable, available environment for consecutive working days;",
        "activities are scheduled in priority order: " + ", ".join(ACTIVITY_PRIORITY) + ".",
        "Go-live preparation starts after readiness (max of application and inflight order tests).",
    ]
    tf.text = "\n".join(lines)
    for p in tf.paragraphs:
        p.runs[0].font.size = Pt(9) if p.runs else 9

def build_deck(plans, sched, path):
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide_exec_summary(prs, plans)
    for name in ACTIVITY_PRIORITY:
        if name in plans:
            slide_activity(prs, plans[name])
    slide_env_utilization(prs, sched)
    slide_assumptions(prs)
    prs.save(path)

# ----------------------------------------------------------------------------
#  TEXT OUTPUT (console)
# ----------------------------------------------------------------------------

def print_plans(plans, sched):
    print("=" * 100)
    print("MIGRATION PLAN  (start %s, %d working days/week)" % (fmt(PLAN_START), WORKDAYS_PER_WEEK))
    print("=" * 100)
    for name in ACTIVITY_PRIORITY:
        if name not in plans:
            continue
        a = plans[name]
        print(f"\n--- {a.name} ({a.type}) ---")
        for ph in PHASES:
            if ph in a.phases:
                s, e = a.phases[ph]
                print(f"  {ph:24s} {fmt(s)} -> {fmt(e)}")
        for run in a.runs:
            print(f"  {run.kind} {run.size:8s} @ {run.env:6s} {fmt(run.start)} -> {fmt(run.end)}")
        print(f"  {'Ready for migration':24s} {fmt(a.ready)}")
        print(f"  {'Go-live prep complete':24s} {fmt(a.golive)}")
    print("\nEnvironment bookings:")
    for e in sorted(sched.env_bookings):
        print(f"  {e}: {len(sched.env_bookings[e])} days booked")

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
    build_deck(plans, sched, args.output)
    print(f"\nPowerPoint deck written to: {args.output}")

if __name__ == "__main__":
    main()

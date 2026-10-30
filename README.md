# projplan40

Migration planning tool: takes planning assumptions (activities, phase durations,
test sizes, environment availability) from parameter tables, builds a working-day
schedule that respects environment capacity constraints, and produces a PowerPoint
deck with different levels of detail.

## Usage

    pip install python-pptx
    python3 migration_plan.py [-o output.pptx]

All planning assumptions are editable parameter tables at the top of
`migration_plan.py` (`ACTIVITIES`, `PHASE_PARAMS`, `TEST_SIZES`, `ENVIRONMENTS`).

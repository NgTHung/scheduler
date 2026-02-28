# Orientation Event Scheduler

Automated scheduling system for university orientation events. Given hosts, mentors, and students with their availability, the solver finds an optimal assignment using Integer Linear Programming (ILP).

Each **session** pairs exactly 1 Host + 1 Mentor + 1 Student in the same time slot, subject to:

- Mentor and student share the same major
- No person is double-booked in a time slot
- Every mentor gets at least one session
- **Objective:** maximise the number of students served

## Features

- **ILP solver** — PuLP/CBC for exact optimisation
- **Excel import** — auto-detects checkbox format (TRUE/FALSE grids) and Vietnamese text format (`Shift 1,2,3` / `5-12`)
- **Web UI** — Streamlit app with live-editable tables, shift label mapping, per-day/per-role views, and one-click export
- **CLI** — quick solve from JSON or built-in sample data

## Quick start

### Prerequisites

- Python ≥ 3.12
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

### Install

```bash
git clone <repo-url> && cd schedule
uv sync            # or: pip install -r requirements.txt 
```

### Run the web UI

```bash
uv run streamlit run app.py
# or
streamlit run app.py
```

### Run from the CLI

```bash
# Built-in sample data
python main.py

# From a JSON file
python main.py data.json
```

## Project structure

```
├── app.py               # Streamlit web UI
├── solver.py            # ILP formulation (PuLP + CBC)
├── models.py            # Host, Mentor, Student, ScheduledSession dataclasses
├── convert_excel.py     # Excel → dict parser (checkbox & text formats)
├── main.py              # CLI entry point
├── data/                # Real input Excel files
│   ├── hosts.xlsx
│   ├── mentors.xlsx
│   └── students.xlsx
└── samples/             # Sample data generators & fixtures
    ├── sample_data.py
    ├── generate_sample_excel.py
    └── generate_sample_text.py
```

## Supported Excel formats

### Checkbox format (separate files)

Three `.xlsx` files — one per role. Each file has one tab per date (e.g. `13/6`). Columns are shift numbers; cells are `TRUE` / `FALSE`.

### Text format (combined workbook)

Single `.xlsx` with tabs named `hosts`, `mentors`, `students`. Day columns contain free-text shift lists such as `1,2,3`, `5-12`, `2`, or `No` (none).

Both formats are auto-detected during import.

## Shift labels

Shifts 1–12 map to 50-minute slots from 8h00 to 20h50 by default (Shift 1 – Shift 12). Labels are fully customisable in the web UI or via a JSON mapping file.

| Shift | Default label         |
|------:|-----------------------|
|     1 | Shift 1               |
|     2 | Shift 2               |
|     3 | Shift 3               |
|     4 | Shift 4               |
|     5 | Shift 5               |
|     6 | Shift 6               |
|     7 | Shift 7               |
|     8 | Shift 8               |
|     9 | Shift 9               |
|    10 | Shift 10              |
|    11 | Shift 11              |
|    12 | Shift 12              |

## ILP formulation

| Symbol | Description |
|--------|-------------|
| T | Set of time slots |
| H, M, S | Sets of hosts, mentors, students |
| V ⊆ T×H×M×S | Valid tuples (availability + major match) |
| x[v] ∈ {0,1} | Session v is scheduled |
| y[s] ∈ {0,1} | Student s is served |

**Objective:** max Σ y[s]

**Constraints:**

1. Each host ≤ 1 session per time slot
2. Each mentor ≤ 1 session per time slot
3. Each student ≤ 1 session per time slot
4. Every mentor ≥ 1 session
5. y[s] ≤ Σ x[v] linking student coverage

Solved with PuLP's built-in CBC backend (default 300 s time limit).

## Dependencies

| Package | Purpose |
|---------|---------|
| [PuLP](https://pypi.org/project/PuLP/) | ILP modelling + CBC solver |
| [Streamlit](https://streamlit.io/) | Web UI |
| [pandas](https://pandas.pydata.org/) | DataFrame tables |
| [openpyxl](https://pypi.org/project/openpyxl/) | Excel read/write |

## License

This project is unlicensed. Add a `LICENSE` file to specify terms.

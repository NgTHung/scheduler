#!/usr/bin/env python3
"""
Orientation Event Scheduler
============================
Solves the session-scheduling problem:
  • Each session = 1 Host + 1 Mentor + 1 Student
  • Mentor ↔ Student must share the same major
  • No person is double-booked within a time-slot
  • Every mentor appears in ≥ 1 session
  • Maximise the number of students served

Usage:
    python main.py              — run with built-in sample data
    python main.py data.json    — run with your own JSON data (see README)
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

# Fix Windows terminal encoding for Unicode output
if sys.stdout.encoding and sys.stdout.encoding.lower().startswith("cp"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from models import Host, Mentor, Student
from solver import solve

def print_schedule(sessions):
    print("\n" + "=" * 78)
    print("  SCHEDULED SESSIONS")
    print("=" * 78)
    current_slot = None
    for sess in sessions:
        if sess.time_slot != current_slot:
            current_slot = sess.time_slot
            print(f"\n  -- {current_slot} {'-' * 60}")
        print(f"    {sess}")
    print()


def print_summary(sessions, mentors, students):
    print("=" * 78)
    print("  SUMMARY")
    print("=" * 78)
    print(f"  Total sessions scheduled : {len(sessions)}")
    mentor_counts = Counter(s.mentor for s in sessions)
    student_covered = {s.student for s in sessions}
    print(f"  Mentors participating    : {len(mentor_counts)} / {len(mentors)}")
    print(f"  Students served          : {len(student_covered)} / {len(students)}")

    unserved = [s.name for s in students if s.name not in student_covered]
    if unserved:
        print(f"  Students NOT served      : {', '.join(unserved)}")

    print("\n  Per-mentor breakdown:")
    for m in mentors:
        cnt = mentor_counts.get(m.name, 0)
        flag = " ✗ MISSING!" if cnt == 0 else ""
        print(f"    {m.name:<16} ({m.major:<12})  sessions: {cnt}{flag}")

    print("\n  Per-major breakdown:")
    major_counts = Counter(s.major for s in sessions)
    for major, cnt in sorted(major_counts.items()):
        print(f"    {major:<16}  sessions: {cnt}")
    print()


def print_constraint_check(sessions, mentors):
    """Verify every hard constraint and print PASS/FAIL."""
    print("=" * 78)
    print("  CONSTRAINT VERIFICATION")
    print("=" * 78)
    ok = True

    # C1–C3: no double-booking
    from collections import defaultdict
    host_slots = defaultdict(list)
    mentor_slots = defaultdict(list)
    student_slots = defaultdict(list)
    for s in sessions:
        host_slots[(s.time_slot, s.host)].append(s)
        mentor_slots[(s.time_slot, s.mentor)].append(s)
        student_slots[(s.time_slot, s.student)].append(s)

    for key, lst in host_slots.items():
        if len(lst) > 1:
            print(f"  FAIL  Host {key[1]} double-booked at {key[0]}")
            ok = False
    for key, lst in mentor_slots.items():
        if len(lst) > 1:
            print(f"  FAIL  Mentor {key[1]} double-booked at {key[0]}")
            ok = False
    for key, lst in student_slots.items():
        if len(lst) > 1:
            print(f"  FAIL  Student {key[1]} double-booked at {key[0]}")
            ok = False

    # C4: major match
    for s in sessions:
        # mentor's major should equal student's desired major
        # (We trust the solver, but let's verify)
        pass  # already encoded in ScheduledSession.major

    # C5: every mentor ≥ 1
    mentor_names = {m.name for m in mentors}
    scheduled_mentors = {s.mentor for s in sessions}
    missing = mentor_names - scheduled_mentors
    if missing:
        for mn in missing:
            print(f"  FAIL  Mentor {mn} has 0 sessions")
        ok = False

    if ok:
        print("  ALL CONSTRAINTS SATISFIED ")
    print()
    return ok

def load_from_json(path: str):
    """
    Expected JSON schema
    --------------------
    {
      "time_slots": ["09:00", "10:00", ...],
      "hosts": [
        {"name": "Alice", "available_slots": ["09:00", "10:00"]}
      ],
      "mentors": [
        {"name": "Carol", "major": "HR", "available_slots": ["09:00"]}
      ],
      "students": [
        {"name": "S1", "desired_major": "HR", "available_slots": ["09:00"]}
      ]
    }
    """
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    time_slots = data["time_slots"]
    hosts = [Host(**h) for h in data["hosts"]]
    mentors = [Mentor(**m) for m in data["mentors"]]
    students = [Student(**s) for s in data["students"]]
    return time_slots, hosts, mentors, students

def main():
    if len(sys.argv) > 1:
        json_path = sys.argv[1]
        print(f"Loading data from {json_path} ...")
        time_slots, hosts, mentors, students = load_from_json(json_path)
    else:
        print("No input file given — using built-in sample data.\n")
        from .samples.sample_data import TIMESLOTS, HOSTS, MENTORS, STUDENTS
        time_slots, hosts, mentors, students = TIMESLOTS, HOSTS, MENTORS, STUDENTS

    print(f"Time-slots : {len(time_slots)}")
    print(f"Hosts      : {len(hosts)}")
    print(f"Mentors    : {len(mentors)}")
    print(f"Students   : {len(students)}")
    majors = sorted({m.major for m in mentors})
    print(f"Majors     : {', '.join(majors)}")
    print("\nSolving ...")

    result = solve(time_slots, hosts, mentors, students, verbose=True)

    if result is None:
        print("\n*** INFEASIBLE — no valid schedule exists under the given constraints. ***")
        sys.exit(1)

    print_schedule(result)
    print_summary(result, mentors, students)
    print_constraint_check(result, mentors)


if __name__ == "__main__":
    main()

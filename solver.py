"""
ILP-based solver for the orientation session scheduling problem.

Formulation
-----------
Sets
    T = time-slots,  H = hosts,  M = mentors,  S = students
    V ⊆ TxHxMxS  — *valid* session tuples where
        • host  h is free at t
        • mentor m is free at t
        • student s is free at t
        • major(m) == desired_major(s)

Decision variables
    x[v] ∈ {0, 1}  for each v ∈ V      — session v is scheduled

    y[s] ∈ {0, 1}  for each student s   — student s is served
                                           (appears in ≥ 1 session)

Objective
    maximise  Σ_s  y[s]                 — serve as many students as possible

Hard constraints
    C1  ∀ t ∈ T, h ∈ H :  Σ_{(t,h,*,*) ∈ V}  x[v] ≤ 1
        (a host is in at most one session per time-slot)

    C2  ∀ t ∈ T, m ∈ M :  Σ_{(t,*,m,*) ∈ V}  x[v] ≤ 1
        (a mentor is in at most one session per time-slot)

    C3  ∀ t ∈ T, s ∈ S :  Σ_{(t,*,*,s) ∈ V}  x[v] ≤ 1
        (a student is in at most one session per time-slot)

    C4  ∀ m ∈ M :  Σ_{(*,*,m,*) ∈ V}  x[v] ≥ 1
        (every mentor must be scheduled at least once)

    C5  ∀ s ∈ S :  y[s]  ≤  Σ_{(*,*,*,s) ∈ V}  x[v]
        (link coverage indicator to actual assignment)
"""

from __future__ import annotations

from collections import defaultdict
from typing import Optional

from pulp import (
    LpProblem,
    LpMaximize,
    LpVariable,
    lpSum,
    PULP_CBC_CMD,
    LpStatusOptimal,
)

from models import Host, Mentor, Student, ScheduledSession


# --------------------------------------------------------------------------- #
#  Index builder — avoids repeated O(|V|) scans when posting constraints       #
# --------------------------------------------------------------------------- #

def _build_indices(valid_sessions: list[tuple[str, str, str, str]]):
    """Return dicts mapping (time, person) → list of session indices."""
    by_host_time: dict[tuple[str, str], list[int]] = defaultdict(list)
    by_mentor_time: dict[tuple[str, str], list[int]] = defaultdict(list)
    by_student_time: dict[tuple[str, str], list[int]] = defaultdict(list)
    by_mentor: dict[str, list[int]] = defaultdict(list)
    by_student: dict[str, list[int]] = defaultdict(list)

    for i, (t, h, m, s) in enumerate(valid_sessions):
        by_host_time[(t, h)].append(i)
        by_mentor_time[(t, m)].append(i)
        by_student_time[(t, s)].append(i)
        by_mentor[m].append(i)
        by_student[s].append(i)

    return by_host_time, by_mentor_time, by_student_time, by_mentor, by_student


# --------------------------------------------------------------------------- #
#  Main solver                                                                 #
# --------------------------------------------------------------------------- #

def solve(
    time_slots: list[str],
    hosts: list[Host],
    mentors: list[Mentor],
    students: list[Student],
    *,
    time_limit_sec: int = 300,
    verbose: bool = False,
) -> Optional[list[ScheduledSession]]:
    """
    Solve the scheduling problem.

    Returns
    -------
    list[ScheduledSession]  on success (may be empty if nothing needed).
    None                    if the problem is infeasible.
    """

    # ---- build mentor-major lookup ---------------------------------------- #
    mentor_major: dict[str, str] = {m.name: m.major for m in mentors}

    # ---- enumerate valid session tuples ----------------------------------- #
    host_avail = {h.name: set(h.available_slots) for h in hosts}
    mentor_avail = {m.name: set(m.available_slots) for m in mentors}
    student_avail = {s.name: set(s.available_slots) for s in students}
    student_major = {s.name: s.desired_major for s in students}

    valid_sessions: list[tuple[str, str, str, str]] = []  # (t, h, m, s)
    for t in time_slots:
        free_hosts = [h.name for h in hosts if t in host_avail[h.name]]
        free_mentors = [m.name for m in mentors if t in mentor_avail[m.name]]
        free_students = [s.name for s in students if t in student_avail[s.name]]
        for h in free_hosts:
            for m in free_mentors:
                for s in free_students:
                    if mentor_major[m] == student_major[s]:
                        valid_sessions.append((t, h, m, s))

    if verbose:
        print(f"  Valid session candidates: {len(valid_sessions)}")

    # ---- quick infeasibility check ---------------------------------------- #
    mentors_with_options = {m.name for m in mentors}
    mentors_in_valid = {m for (_, _, m, _) in valid_sessions}
    impossible_mentors = mentors_with_options - mentors_in_valid
    if impossible_mentors:
        print(
            "INFEASIBLE — the following mentors have NO valid session "
            "(no student wants their major, or schedules don't overlap):"
        )
        for mn in sorted(impossible_mentors):
            print(f"  • {mn} ({mentor_major[mn]})")
        return None

    # ---- build index structures ------------------------------------------- #
    (by_host_time, by_mentor_time, by_student_time,
     by_mentor, by_student) = _build_indices(valid_sessions)

    # ---- create ILP ------------------------------------------------------- #
    prob = LpProblem("OrientationScheduling", LpMaximize)

    # x[i] — session i is scheduled
    x = [LpVariable(f"x_{i}", cat="Binary") for i in range(len(valid_sessions))]

    # y[s] — student s is served (coverage indicator)
    y = {s.name: LpVariable(f"y_{s.name}", cat="Binary") for s in students}

    # ---- objective: maximise student coverage ----------------------------- #
    prob += lpSum(y[s.name] for s in students), "MaxStudentsCovered"

    # ---- C1: host ≤ 1 session per time-slot ------------------------------- #
    for (t, h), idxs in by_host_time.items():
        prob += lpSum(x[i] for i in idxs) <= 1, f"HostSlot_{h}_{t}"

    # ---- C2: mentor ≤ 1 session per time-slot ----------------------------- #
    for (t, m), idxs in by_mentor_time.items():
        prob += lpSum(x[i] for i in idxs) <= 1, f"MentorSlot_{m}_{t}"

    # ---- C3: student ≤ 1 session per time-slot ---------------------------- #
    for (t, s), idxs in by_student_time.items():
        prob += lpSum(x[i] for i in idxs) <= 1, f"StudentSlot_{s}_{t}"

    # ---- C4: every mentor in ≥ 1 session ---------------------------------- #
    for m in mentors:
        idxs = by_mentor[m.name]
        prob += lpSum(x[i] for i in idxs) >= 1, f"MentorMin_{m.name}"

    # ---- C5: link y[s] to assignments ------------------------------------- #
    for s in students:
        idxs = by_student.get(s.name, [])
        if idxs:
            prob += y[s.name] <= lpSum(x[i] for i in idxs), f"Link_{s.name}"
        else:
            prob += y[s.name] == 0, f"Link_{s.name}"

    # ---- solve ------------------------------------------------------------ #
    solver = PULP_CBC_CMD(msg=int(verbose), timeLimit=time_limit_sec)
    prob.solve(solver)

    if prob.status != LpStatusOptimal:
        return None

    # ---- extract solution ------------------------------------------------- #
    scheduled: list[ScheduledSession] = []
    for i, (t, h, m, s) in enumerate(valid_sessions):
        if x[i].varValue is not None and x[i].varValue > 0.5:
            scheduled.append(
                ScheduledSession(
                    time_slot=t,
                    host=h,
                    mentor=m,
                    student=s,
                    major=mentor_major[m],
                )
            )

    scheduled.sort(key=lambda sess: (sess.time_slot, sess.major))
    return scheduled

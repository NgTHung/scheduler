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

    # ---- merge duplicate mentors (same name, different majors/slots) ------ #
    merged: dict[str, Mentor] = {}
    for m in mentors:
        if m.name in merged:
            existing = merged[m.name]
            # Combine majors (comma-separated) and slots
            existing_majors = set(existing.majors)
            new_majors = set(m.majors)
            combined = existing_majors | new_majors
            existing.major = ", ".join(sorted(combined))
            existing.available_slots = list(
                dict.fromkeys(existing.available_slots + m.available_slots)
            )
        else:
            # Copy so we don't mutate the original
            merged[m.name] = Mentor(
                name=m.name,
                major=m.major,
                available_slots=list(m.available_slots),
            )
    mentors = list(merged.values())

    # ---- build mentor-major lookup ---------------------------------------- #
    mentor_major: dict[str, str] = {m.name: m.major for m in mentors}

    # ---- normalise major names for matching ------------------------------- #
    def _norm_major(s: str) -> str:
        return s.strip().lower()

    # Build a set of normalised majors per mentor (supports multi-major)
    mentor_norm_majors: dict[str, set[str]] = {
        m.name: {_norm_major(mj) for mj in m.majors} for m in mentors
    }

    # ---- enumerate valid session tuples ----------------------------------- #
    host_avail = {h.name: set(h.available_slots) for h in hosts}
    mentor_avail = {m.name: set(m.available_slots) for m in mentors}
    student_avail = {s.name: set(s.available_slots) for s in students}
    student_major = {s.name: s.desired_major for s in students}

    valid_sessions: list[tuple[str, str, str, str]] = []  # (t, h, m, s)
    seen_sessions: set[tuple[str, str, str, str]] = set()
    for t in time_slots:
        free_hosts = [h.name for h in hosts if t in host_avail[h.name]]
        free_mentors = [m.name for m in mentors if t in mentor_avail[m.name]]
        free_students = [s.name for s in students if t in student_avail[s.name]]
        for h in free_hosts:
            for m in free_mentors:
                if h == m:
                    continue  # same person can't be host and mentor
                for s in free_students:
                    if s == h or s == m:
                        continue  # same person can't fill two roles
                    key = (t, h, m, s)
                    if (
                        key not in seen_sessions
                        and _norm_major(student_major[s]) in mentor_norm_majors[m]
                    ):
                        seen_sessions.add(key)
                        valid_sessions.append(key)

    if verbose:
        print(f"  Valid session candidates: {len(valid_sessions)}")

    # ---- quick infeasibility check ---------------------------------------- #
    mentors_with_options = {m.name for m in mentors}
    mentors_in_valid = {m for (_, _, m, _) in valid_sessions}
    impossible_mentors = mentors_with_options - mentors_in_valid
    if impossible_mentors:
        error = (
            "INFEASIBLE — the following mentors have NO valid session "
            + "(no student wants their major, or schedules don't overlap):"
        )
        for mn in sorted(impossible_mentors):
            error += (f"  • {mn} ({mentor_major[mn]})")
        raise Exception(error)

    # ---- build index structures ------------------------------------------- #
    (by_host_time, by_mentor_time, by_student_time, by_mentor, by_student) = (
        _build_indices(valid_sessions)
    )

    # ---- create ILP ------------------------------------------------------- #
    prob = LpProblem("OrientationScheduling", LpMaximize)

    # x[i] — session i is scheduled
    x = [LpVariable(f"x_{i}", cat="Binary") for i in range(len(valid_sessions))]

    # y[s] — student s is served (coverage indicator)
    y = {s.name: LpVariable(f"y_{i}", cat="Binary") for i, s in enumerate(students)}

    # ---- objective: maximise student coverage ----------------------------- #
    prob += lpSum(y[s.name] for s in students), "MaxStudentsCovered"

    # ---- C1: host ≤ 1 session per time-slot ------------------------------- #
    for ci, ((t, h), idxs) in enumerate(by_host_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C1_{ci}"

    # ---- C2: mentor ≤ 1 session per time-slot ----------------------------- #
    for ci, ((t, m), idxs) in enumerate(by_mentor_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C2_{ci}"

    # ---- C3: student ≤ 1 session per time-slot ---------------------------- #
    for ci, ((t, s), idxs) in enumerate(by_student_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C3_{ci}"

    # ---- C4: every mentor in ≥ 1 session ---------------------------------- #
    for ci, m in enumerate(mentors):
        idxs = by_mentor[m.name]
        prob += lpSum(x[i] for i in idxs) >= 1, f"C4_{ci}"

    # ---- C5: link y[s] to assignments ------------------------------------- #
    for ci, s in enumerate(students):
        idxs = by_student.get(s.name, [])
        if idxs:
            prob += y[s.name] <= lpSum(x[i] for i in idxs), f"C5_{ci}"
        else:
            prob += y[s.name] == 0, f"C5_{ci}"

    # ---- C6: cross-role no-double-booking --------------------------------- #
    # If the same person name appears in multiple roles (e.g. host AND student),
    # they can participate in at most 1 session per time-slot across ALL roles.
    all_names: set[str] = set()
    host_names = {h.name for h in hosts}
    mentor_names = {m.name for m in mentors}
    student_names = {s.name for s in students}
    cross_role = (
        (host_names & mentor_names)
        | (host_names & student_names)
        | (mentor_names & student_names)
    )

    if cross_role:
        # Build per-(time, person) index across all roles
        by_person_time: dict[tuple[str, str], list[int]] = defaultdict(list)
        for i, (t, h, m, s) in enumerate(valid_sessions):
            if h in cross_role:
                by_person_time[(t, h)].append(i)
            if m in cross_role:
                by_person_time[(t, m)].append(i)
            if s in cross_role:
                by_person_time[(t, s)].append(i)

        ci = 0
        for (t, person), idxs in by_person_time.items():
            # Deduplicate indices (a person could be host+student in same tuple)
            unique_idxs = list(dict.fromkeys(idxs))
            prob += lpSum(x[i] for i in unique_idxs) <= 1, f"C6_{ci}"
            ci += 1

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

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

    u[m] ∈ {0, 1}  for each mentor m    — mentor m is used
                                           (appears in ≥ 1 session)

    z[h] ∈ {0, 1}  for each host h      — host h is used
                                           (appears in ≥ 1 session)

    L ∈ Z≥0                            — maximum sessions assigned to any mentor

Objective
    maximise, in priority order:
        1.  Σ_s y[s]                    — serve as many students as possible
        2.  Σ_m u[m]                    — use as many different mentors as possible
        3.  Σ_h z[h]                    — use as many different hosts as possible
        4. -L                           — minimise the busiest mentor's load
        5. -Σ_v x[v]                    — avoid unnecessary sessions

Hard constraints
    C1  ∀ t ∈ T, h ∈ H :  Σ_{(t,h,*,*) ∈ V}  x[v] ≤ 1
        (a host is in at most one session per time-slot)

    C2  ∀ t ∈ T, m ∈ M :  Σ_{(t,*,m,*) ∈ V}  x[v] ≤ 1
        (a mentor is in at most one session per time-slot)

    C3  ∀ t ∈ T, s ∈ S :  Σ_{(t,*,*,s) ∈ V}  x[v] ≤ 1
        (a student is in at most one session per time-slot)

    C4  ∀ s ∈ S :  y[s]  ↔  Σ_{(*,*,*,s) ∈ V}  x[v] ≥ 1
        (link coverage indicator to actual assignments)

    C5  ∀ m ∈ M :  u[m]  ↔  Σ_{(*,*,m,*) ∈ V}  x[v] ≥ 1
        (link mentor-used indicator to actual assignments)

    C6  ∀ h ∈ H :  z[h]  ↔  Σ_{(*,h,*,*) ∈ V}  x[v] ≥ 1
        (link host-used indicator to actual assignments)

    C7  ∀ m ∈ M :  Σ_{(*,*,m,*) ∈ V}  x[v] ≤ L
        (L is the maximum mentor session load)
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
    by_host: dict[str, list[int]] = defaultdict(list)
    by_mentor: dict[str, list[int]] = defaultdict(list)
    by_student: dict[str, list[int]] = defaultdict(list)

    for i, (t, h, m, s) in enumerate(valid_sessions):
        by_host_time[(t, h)].append(i)
        by_mentor_time[(t, m)].append(i)
        by_student_time[(t, s)].append(i)
        by_host[h].append(i)
        by_mentor[m].append(i)
        by_student[s].append(i)

    return by_host_time, by_mentor_time, by_student_time, by_host, by_mentor, by_student


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

    # ---- merge duplicate students (same name, different desired_majors/slots) #
    merged_s: dict[str, Student] = {}
    for s in students:
        if s.name in merged_s:
            existing = merged_s[s.name]
            existing_majors = set(existing.desired_majors)
            new_majors = set(s.desired_majors)
            combined = existing_majors | new_majors
            existing.desired_major = ", ".join(sorted(combined))
            existing.available_slots = list(
                dict.fromkeys(existing.available_slots + s.available_slots)
            )
        else:
            merged_s[s.name] = Student(
                name=s.name,
                desired_major=s.desired_major,
                available_slots=list(s.available_slots),
            )
    students = list(merged_s.values())

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

    # Build a set of normalised desired majors per student (multi-major)
    student_norm_majors: dict[str, set[str]] = {
        s.name: {_norm_major(mj) for mj in s.desired_majors} for s in students
    }

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
                        and mentor_norm_majors[m] & student_norm_majors[s]
                    ):
                        seen_sessions.add(key)
                        valid_sessions.append(key)

    if verbose:
        print(f"  Valid session candidates: {len(valid_sessions)}")

    # ---- build index structures ------------------------------------------- #
    (by_host_time, by_mentor_time, by_student_time, by_host, by_mentor, by_student) = (
        _build_indices(valid_sessions)
    )

    # ---- create ILP ------------------------------------------------------- #
    prob = LpProblem("OrientationScheduling", LpMaximize)

    # x[i] — session i is scheduled
    x = [LpVariable(f"x_{i}", cat="Binary") for i in range(len(valid_sessions))]

    # y[s] — student s is served (coverage indicator)
    y = {s.name: LpVariable(f"y_{i}", cat="Binary") for i, s in enumerate(students)}

    # u[m] — mentor m is used at least once (soft coverage indicator)
    u = {m.name: LpVariable(f"u_{i}", cat="Binary") for i, m in enumerate(mentors)}

    # z[h] — host h is used at least once (soft diversity indicator)
    z = {h.name: LpVariable(f"z_{i}", cat="Binary") for i, h in enumerate(hosts)}

    # L — maximum number of sessions assigned to any mentor
    max_mentor_load = LpVariable("max_mentor_load", lowBound=0, cat="Integer")

    # ---- objective: maximise students, mentors, hosts; minimise load/sessions
    # Five-tier weights:
    #   (1) one extra served student beats any mentor/host/load/session tradeoff,
    #   (2) one extra active mentor beats any host/load/session tradeoff,
    #   (3) one extra distinct host beats any mentor-load/session tradeoff,
    #   (4) lower max mentor load beats any session-count increase,
    #   (5) fewer sessions wins only after students, mentors, hosts, and load tie.
    session_weight = 1
    mentor_load_weight = len(valid_sessions) + 1
    host_weight = len(valid_sessions) * mentor_load_weight + len(valid_sessions) + 1
    mentor_weight = len(hosts) * host_weight + len(valid_sessions) * mentor_load_weight + len(valid_sessions) + 1
    student_weight = len(mentors) * mentor_weight + len(hosts) * host_weight + len(valid_sessions) * mentor_load_weight + len(valid_sessions) + 1
    prob += (
        lpSum(y[s.name] * student_weight for s in students)
        + lpSum(u[m.name] * mentor_weight for m in mentors)
        + lpSum(z[h.name] * host_weight for h in hosts)
        - max_mentor_load * mentor_load_weight
        - lpSum(x[i] * session_weight for i in range(len(valid_sessions)))
    ), "MaxStudentsMaxMentorsMaxHostsBalanceMentorsMinSessions"

    # ---- C1: host ≤ 1 session per time-slot ------------------------------- #
    for ci, ((t, h), idxs) in enumerate(by_host_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C1_{ci}"

    # ---- C2: mentor ≤ 1 session per time-slot ----------------------------- #
    for ci, ((t, m), idxs) in enumerate(by_mentor_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C2_{ci}"

    # ---- C3: student ≤ 1 session per time-slot ---------------------------- #
    for ci, ((t, s), idxs) in enumerate(by_student_time.items()):
        prob += lpSum(x[i] for i in idxs) <= 1, f"C3_{ci}"

    # ---- C4: link y[s] to assignments ------------------------------------- #
    ci4_link = 0
    for ci, s in enumerate(students):
        idxs = by_student.get(s.name, [])
        if idxs:
            prob += y[s.name] <= lpSum(x[i] for i in idxs), f"C4_{ci}"
            for i in idxs:
                prob += x[i] <= y[s.name], f"C4_link_{ci4_link}"
                ci4_link += 1
        else:
            prob += y[s.name] == 0, f"C4_{ci}"

    # ---- C5: link u[m] to mentor assignments ------------------------------ #
    ci5_link = 0
    for ci, m in enumerate(mentors):
        idxs = by_mentor.get(m.name, [])
        if idxs:
            prob += u[m.name] <= lpSum(x[i] for i in idxs), f"C5_{ci}"
            for i in idxs:
                prob += x[i] <= u[m.name], f"C5_link_{ci5_link}"
                ci5_link += 1
        else:
            prob += u[m.name] == 0, f"C5_{ci}"

    # ---- C6: link z[h] to host assignments -------------------------------- #
    ci6_link = 0
    for ci, h in enumerate(hosts):
        idxs = by_host.get(h.name, [])
        if idxs:
            prob += z[h.name] <= lpSum(x[i] for i in idxs), f"C6_{ci}"
            for i in idxs:
                prob += x[i] <= z[h.name], f"C6_link_{ci6_link}"
                ci6_link += 1
        else:
            prob += z[h.name] == 0, f"C6_{ci}"

    # ---- C7: bound max mentor load ---------------------------------------- #
    for ci, m in enumerate(mentors):
        idxs = by_mentor[m.name]
        prob += lpSum(x[i] for i in idxs) <= max_mentor_load, f"C7_{ci}"

    # ---- C8: multi-major students get ≥1 session per desired major -------- #
    # For each (student, desired_major) pair, require at least one session
    # with a mentor covering that major — but only when student is served.
    ci8 = 0
    for s in students:
        majors = student_norm_majors[s.name]
        if len(majors) <= 1:
            continue  # single-major students already handled by C4
        s_idxs = set(by_student.get(s.name, []))
        for mj in majors:
            # Find session indices where this student is paired with a
            # mentor whose majors include this specific desired major.
            matching = [
                i for i in s_idxs
                if mj in mentor_norm_majors[valid_sessions[i][2]]
            ]
            if matching:
                prob += lpSum(x[i] for i in matching) >= y[s.name], f"C8_{ci8}"
            ci8 += 1

    # ---- C9: cross-role no-double-booking --------------------------------- #
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
            prob += lpSum(x[i] for i in unique_idxs) <= 1, f"C9_{ci}"
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
            # Determine the specific major for this session by intersecting
            # the mentor's majors with the student's desired majors.
            matched = mentor_norm_majors[m] & student_norm_majors[s]
            if matched:
                # Pick the original-cased form from the mentor's major list
                norm_to_orig = {_norm_major(mj): mj for mj in
                                merged[m].major.replace("|", ",").replace(";", ",").replace("/", ",").split(",")}
                session_major = next(
                    (norm_to_orig[nm] for nm in matched if nm in norm_to_orig),
                    mentor_major[m],
                )
            else:
                session_major = mentor_major[m]
            scheduled.append(
                ScheduledSession(
                    time_slot=t,
                    host=h,
                    mentor=m,
                    student=s,
                    major=session_major.strip(),
                )
            )

    scheduled.sort(key=lambda sess: (sess.time_slot, sess.major))
    return scheduled

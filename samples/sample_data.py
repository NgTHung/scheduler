"""
Sample data for demonstration / testing.

Scenario
--------
  Time-slots : 09:00, 10:00, 11:00, 14:00, 15:00
  Majors     : HR, Sales, Marketing
  Hosts      : 2
  Mentors    : 5  (2 HR, 2 Sales, 1 Marketing)
  Students   : 8  (various desired majors)

The schedules are intentionally *tight* so the solver actually has
something non-trivial to figure out.
"""

from models import Host, Mentor, Student

TIMESLOTS = ["09:00", "10:00", "11:00", "14:00", "15:00"]

HOSTS = [
    Host("Alice_H",   ["09:00", "10:00", "11:00", "14:00", "15:00"]),
    Host("Bob_H",     ["09:00", "10:00", "14:00", "15:00"]),
]

MENTORS = [
    Mentor("Carol_M",   "HR",        ["09:00", "10:00", "11:00"]),
    Mentor("Dave_M",    "HR",        ["14:00", "15:00"]),
    Mentor("Eve_M",     "Sales",     ["09:00", "10:00", "14:00"]),
    Mentor("Frank_M",   "Sales",     ["10:00", "11:00", "15:00"]),
    Mentor("Grace_M",   "Marketing", ["09:00", "11:00", "14:00", "15:00"]),
]

STUDENTS = [
    Student("S1_Liam",    "HR",        ["09:00", "10:00"]),
    Student("S2_Noah",    "HR",        ["10:00", "14:00", "15:00"]),
    Student("S3_Emma",    "Sales",     ["09:00", "10:00", "11:00"]),
    Student("S4_Olivia",  "Sales",     ["10:00", "14:00"]),
    Student("S5_Ava",     "Sales",     ["11:00", "15:00"]),
    Student("S6_Sophia",  "Marketing", ["09:00", "11:00"]),
    Student("S7_Mia",     "Marketing", ["14:00", "15:00"]),
    Student("S8_James",   "HR",        ["11:00", "14:00"]),
]

"""
Data models for the orientation session scheduling problem.

Problem:
  - Each SESSION consists of exactly 3 people: 1 Host, 1 Mentor, 1 Student.
  - Every person has a set of time-slots they are free.
  - Mentors have a major (e.g. HR, Sales, Marketing).
  - Students have a desired major — they must be paired with a mentor of
    that same major.
  - Hosts are NOT restricted by major and may participate in as many
    sessions as needed (but only one session per time-slot).
  - Every mentor must appear in AT LEAST one session.
"""

from __future__ import annotations
from dataclasses import dataclass, field


@dataclass
class Host:
    name: str
    available_slots: list[str] = field(default_factory=list)

    def __repr__(self) -> str:
        return f"Host({self.name})"


@dataclass
class Mentor:
    name: str
    major: str
    available_slots: list[str] = field(default_factory=list)

    def __repr__(self) -> str:
        return f"Mentor({self.name}, {self.major})"


@dataclass
class Student:
    name: str
    desired_major: str
    available_slots: list[str] = field(default_factory=list)

    def __repr__(self) -> str:
        return f"Student({self.name}, wants={self.desired_major})"


@dataclass
class ScheduledSession:
    """One scheduled orientation session."""
    time_slot: str
    host: str
    mentor: str
    student: str
    major: str

    def __repr__(self) -> str:
        return (
            f"[{self.time_slot}]  Host: {self.host:<12}  "
            f"Mentor: {self.mentor:<12} ({self.major})  "
            f"Student: {self.student}"
        )

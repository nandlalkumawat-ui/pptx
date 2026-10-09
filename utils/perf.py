"""
Lightweight stage timing and counters for the generation pipeline.

Uses the same "[timing]" print style as the rest of the project so the
output stays familiar. A single module-level ``PERF`` report collects
stage durations and counters for one run and prints a summary at the end.
"""

import threading
import time
from collections import defaultdict
from contextlib import contextmanager


class PerfReport:

    def __init__(self):
        self._lock = threading.Lock()
        self.reset()

    def reset(self) -> None:
        with self._lock:
            self.timings: dict[str, float] = defaultdict(float)
            self.counters: dict[str, int] = defaultdict(int)
            self._order: list[str] = []

    def add_time(self, stage: str, seconds: float) -> None:
        with self._lock:
            if stage not in self._order:
                self._order.append(stage)
            self.timings[stage] += seconds

    def incr(self, name: str, amount: int = 1) -> None:
        with self._lock:
            self.counters[name] += amount

    @contextmanager
    def stage(self, name: str, verbose: bool = True):
        """Time a block with time.perf_counter() and record it."""

        started = time.perf_counter()
        try:
            yield
        finally:
            elapsed = time.perf_counter() - started
            self.add_time(name, elapsed)
            if verbose:
                print(f"[timing] {name}: {elapsed:.2f}s")

    def summary(self) -> str:
        lines = ["", "========== PERFORMANCE SUMMARY =========="]
        for stage in self._order:
            lines.append(f"  {stage:<34} {self.timings[stage]:>8.2f}s")
        if self.counters:
            lines.append("  ---------------------------------------")
            for name in sorted(self.counters):
                lines.append(f"  {name:<34} {self.counters[name]:>8}")
        lines.append("=========================================")
        return "\n".join(lines)


PERF = PerfReport()

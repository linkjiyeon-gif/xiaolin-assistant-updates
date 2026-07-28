from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from enum import Enum


class TaskPhase(str, Enum):
    IDLE = "idle"
    STARTING = "starting"
    RUNNING = "running"
    STOPPING = "stopping"
    SUCCEEDED = "succeeded"
    FAILED = "failed"


BUSY_PHASES = frozenset({TaskPhase.STARTING, TaskPhase.RUNNING, TaskPhase.STOPPING})


@dataclass(frozen=True)
class TaskSnapshot:
    name: str
    phase: TaskPhase
    message: str
    generation: int
    updated_at: float

    @property
    def is_busy(self) -> bool:
        return self.phase in BUSY_PHASES

    @property
    def is_running(self) -> bool:
        return self.phase == TaskPhase.RUNNING


class RuntimeTaskState:
    """Thread-safe runtime state shared by workers and cached UI pages."""

    def __init__(self, name: str):
        self.name = name
        self._lock = threading.RLock()
        self._phase = TaskPhase.IDLE
        self._message = ""
        self._generation = 0
        self._updated_at = time.monotonic()

    def set(self, phase: TaskPhase | str, message: str = "") -> TaskSnapshot:
        normalized = phase if isinstance(phase, TaskPhase) else TaskPhase(str(phase))
        with self._lock:
            if normalized != self._phase or message != self._message:
                self._generation += 1
                self._phase = normalized
                self._message = str(message or "")
                self._updated_at = time.monotonic()
            return self._snapshot_unlocked()

    def snapshot(self) -> TaskSnapshot:
        with self._lock:
            return self._snapshot_unlocked()

    def _snapshot_unlocked(self) -> TaskSnapshot:
        return TaskSnapshot(
            name=self.name,
            phase=self._phase,
            message=self._message,
            generation=self._generation,
            updated_at=self._updated_at,
        )

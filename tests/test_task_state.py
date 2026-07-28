from __future__ import annotations

import threading
import unittest

from app.task_state import RuntimeTaskState, TaskPhase


class RuntimeTaskStateTests(unittest.TestCase):
    def test_phase_and_busy_flags_are_consistent(self):
        state = RuntimeTaskState("network")
        self.assertEqual(TaskPhase.IDLE, state.snapshot().phase)
        self.assertFalse(state.snapshot().is_busy)

        starting = state.set(TaskPhase.STARTING, "initializing")
        self.assertTrue(starting.is_busy)
        self.assertFalse(starting.is_running)

        running = state.set(TaskPhase.RUNNING, "ready")
        self.assertTrue(running.is_busy)
        self.assertTrue(running.is_running)

        stopped = state.set(TaskPhase.IDLE)
        self.assertFalse(stopped.is_busy)
        self.assertGreater(stopped.generation, starting.generation)

    def test_concurrent_reads_never_observe_an_invalid_phase(self):
        state = RuntimeTaskState("adb_tools")
        observed = []

        def writer():
            for phase in (TaskPhase.STARTING, TaskPhase.RUNNING, TaskPhase.STOPPING, TaskPhase.IDLE):
                state.set(phase)

        threads = [threading.Thread(target=writer) for _ in range(4)]
        for thread in threads:
            thread.start()
        while any(thread.is_alive() for thread in threads):
            observed.append(state.snapshot().phase)
        for thread in threads:
            thread.join()

        self.assertTrue(all(isinstance(phase, TaskPhase) for phase in observed))


if __name__ == "__main__":
    unittest.main()

"""A first-come, first-served queue for runs, shared by every user of the app.

Streamlit runs each user's session in the same server process, so this one object sees all of
them. A few runs go at once (they share Companies House's pace either way); anyone else waits
their turn and is told their place in the line.
"""

import itertools
import threading

MAX_ACTIVE_RUNS = 3


class RunQueue:
    def __init__(self, max_active=MAX_ACTIVE_RUNS):
        self.max_active = max_active
        self.cond = threading.Condition()
        self.waiting = []            # tickets, in the order they arrived
        self.active = set()
        self._numbers = itertools.count(1)

    def join(self):
        with self.cond:
            ticket = next(self._numbers)
            self.waiting.append(ticket)
            self._promote()
            return ticket

    def _promote(self):
        while self.waiting and len(self.active) < self.max_active:
            self.active.add(self.waiting.pop(0))
        self.cond.notify_all()

    def position(self, ticket):
        """0 when the run may go; otherwise its place in the line (1 = next)."""
        with self.cond:
            return 0 if ticket in self.active else self.waiting.index(ticket) + 1

    def wait_turn(self, ticket, on_wait=None, poll=2.0):
        """Block until the ticket may run, calling on_wait(position, running) while it waits.
        on_wait is how the page shows the place in the line, and how Streamlit stops a
        waiting run if the user leaves (it raises inside the page's call)."""
        with self.cond:
            while ticket not in self.active:
                if on_wait:
                    self.cond.release()
                    try:
                        on_wait(self.waiting.index(ticket) + 1, len(self.active))
                    finally:
                        self.cond.acquire()
                if ticket not in self.active:
                    self.cond.wait(poll)

    def leave(self, ticket):
        """Done, or gave up waiting: free the place for the next in line."""
        with self.cond:
            self.active.discard(ticket)
            if ticket in self.waiting:
                self.waiting.remove(ticket)
            self._promote()

    def status(self):
        with self.cond:
            return len(self.active), len(self.waiting)


QUEUE = RunQueue()

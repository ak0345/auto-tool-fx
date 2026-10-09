"""Sharing the app between many users: one Companies House pace, a queue for runs, and one
scanned PDF page read at a time."""

import threading
import time

import turnover as t
from run_queue import RunQueue


# ---------------------------------------------------------------- shared pace


def test_every_user_shares_one_pace_by_default():
    assert t.Web().pace is t.Web().pace is t.PACE


def test_pace_spaces_requests_from_different_users(monkeypatch):
    pace = t.Pace(delay=0.2)
    stamps = []

    def user():
        for _ in range(3):
            pace.wait_turn()
            stamps.append(time.monotonic())

    threads = [threading.Thread(target=user) for _ in range(3)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()
    stamps.sort()
    gaps = [b - a for a, b in zip(stamps, stamps[1:])]
    assert len(stamps) == 9 and min(gaps) >= 0.19                      # 3 users, never closer than the pace


def test_push_back_slows_and_pauses_then_recovers():
    pace = t.Pace(delay=1.0)
    pace.pushed_back(cool_down=60)
    assert pace.delay == 2.0 and pace.next_at - time.monotonic() > 59
    for _ in range(3):
        pace.pushed_back(cool_down=0)
    assert pace.delay == t.MAX_DELAY
    for _ in range(500):
        pace.went_through()
    assert pace.delay == 1.0


# ---------------------------------------------------------------- run queue


def test_queue_lets_a_few_run_and_lines_up_the_rest():
    q = RunQueue(max_active=2)
    a, b, c, d = (q.join() for _ in range(4))
    assert [q.position(x) for x in (a, b, c, d)] == [0, 0, 1, 2]
    assert q.status() == (2, 2)
    q.leave(a)                                                         # a finishes: c goes, d moves up
    assert [q.position(x) for x in (b, c, d)] == [0, 0, 1]
    q.leave(d)                                                         # d gives up waiting
    assert q.status() == (2, 0)


def test_waiting_run_starts_when_a_place_frees():
    q = RunQueue(max_active=1)
    first, second = q.join(), q.join()
    seen, started = [], threading.Event()

    def wait():
        q.wait_turn(second, on_wait=lambda place, active: seen.append((place, active)), poll=0.05)
        started.set()

    th = threading.Thread(target=wait)
    th.start()
    time.sleep(0.2)
    assert not started.is_set() and seen[0] == (1, 1)                  # told: 1st in line, 1 running
    q.leave(first)
    assert started.wait(2)
    th.join()


def test_leaving_while_waiting_frees_the_place():
    q = RunQueue(max_active=1)
    first, second, third = q.join(), q.join(), q.join()

    def stop(place, active):
        raise KeyboardInterrupt                                       # what Streamlit does when a user leaves

    try:
        q.wait_turn(second, on_wait=stop)
    except KeyboardInterrupt:
        q.leave(second)
    assert q.position(third) == 1
    q.leave(first)
    assert q.position(third) == 0


# ---------------------------------------------------------------- one scanned page at a time


def test_scanned_pages_are_read_one_at_a_time():
    inside, most = [0], [0]
    lock = threading.Lock()

    def user():
        for _ in range(5):
            t.ocr_turn()
            try:
                with lock:
                    inside[0] += 1
                    most[0] = max(most[0], inside[0])
                time.sleep(0.01)
                with lock:
                    inside[0] -= 1
            finally:
                t.OCR_SLOTS.release()

    threads = [threading.Thread(target=user) for _ in range(4)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()
    assert most[0] == 1


def test_waiting_for_the_reader_is_reported():
    messages = []
    t.OCR_SLOTS.acquire()                                              # someone else is reading
    th = threading.Thread(target=lambda: (t.ocr_turn(messages.append), t.OCR_SLOTS.release()))
    th.start()
    time.sleep(0.1)
    t.OCR_SLOTS.release()
    th.join(2)
    assert messages and "waiting for the scanned-PDF reader" in messages[0]

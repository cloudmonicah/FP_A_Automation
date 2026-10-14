import pytest

from automations.forecast.tableau_service import wait_for_tableau_workbook


class FakeClock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


def test_waits_until_workbook_is_ready():
    clock = FakeClock()
    checks = []

    def find(path):
        checks.append(clock.now)
        if clock.now < 60:  # a large workbook still loading
            raise RuntimeError("not ready yet")
        return "window"

    window = wait_for_tableau_workbook(
        "C:/x/FY27 FORECAST.twb", find_window=find, sleep=clock.sleep, clock=clock
    )
    assert window == "window"
    assert len(checks) > 3  # old code gave up after 3 one-second checks


def test_gives_a_clear_message_when_tableau_never_opens():
    clock = FakeClock()

    def find(path):
        raise RuntimeError("no Tableau window")

    with pytest.raises(TimeoutError) as exc:
        wait_for_tableau_workbook(
            "C:/x/Legal.twb", timeout_seconds=30, find_window=find,
            sleep=clock.sleep, clock=clock,
        )
    assert '"Legal"' in str(exc.value) and "30 seconds" in str(exc.value)

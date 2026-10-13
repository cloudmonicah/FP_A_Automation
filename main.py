from __future__ import annotations

import logging


def main() -> None:
    from shared.logging_setup import configure

    configure()
    log = logging.getLogger("main")
    try:
        from app import FinanceAutomationApp

        FinanceAutomationApp().run()
    except Exception:
        log.exception("The app failed to start")
        raise


if __name__ == "__main__":
    main()

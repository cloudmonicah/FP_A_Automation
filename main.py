from __future__ import annotations

import logging


def _show_startup_error(message: str) -> None:
    """Show a dialog too, because the windowed .exe has no terminal."""
    try:
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("FP&A Automation could not start", message, parent=root)
        root.destroy()
    except Exception:
        pass


def main() -> None:
    from shared.logging_setup import configure, log_file

    configure()
    log = logging.getLogger("main")
    try:
        from app import FinanceAutomationApp

        FinanceAutomationApp().run()
    except ModuleNotFoundError as exc:
        log.exception("The app failed to start")
        _show_startup_error(
            f"A required package is missing: {exc.name}\n\n"
            "Run:  pip install -r requirements.txt\n\n"
            f"Details: {log_file()}"
        )
        raise SystemExit(1) from exc
    except Exception as exc:
        log.exception("The app failed to start")
        _show_startup_error(f"{exc}\n\nDetails: {log_file()}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()

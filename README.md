# FP&A Automation

A desktop app (ttkbootstrap) that automates the monthly FP&A workflow:

- **Monthly Forecast**: refreshes the forecast workbooks from Tableau and validates the result.
- **Actuals Validation**: compares OPEX, Direct Cost and Revenue actuals between Tableau and the financial statements.
- **Initial Budget Setup**: builds the budget workbooks.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

The Tableau automation needs Windows with Tableau Desktop installed.

Copy `settings.example.json` to `settings.json` and set the shared-drive folders,
or set them later on the app's Settings page.

## Run the app

```powershell
python main.py
```

## Tests

```powershell
pip install -r requirements-dev.txt
python -m pytest
```

`tests/manual/` holds older check scripts that need real workbooks and Tableau; pytest skips them.

## Building the .exe for colleagues

Run `build_exe.bat` on Windows. See [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) for what colleagues need,
how to release a version, settings, and where the log file is.

## Project layout

| Path | Contents |
| --- | --- |
| `app.py`, `main.py`, `config.py` | App shell |
| `ui/` | One view per sidebar page |
| `automations/` | Business logic for `forecast`, `actuals` and `budget_setup` |
| `shared/` | Settings, logging and helpers shared across automations |
| `scripts/` | Command-line tools you run by hand (see below) |
| `tests/` | Automated tests (`tests/manual/`: scripts needing real data) |
| `docs/` | Deployment guide |

## Command-line scripts

Run these from the project root as modules so the `automations` package can be found:

```powershell
python -m scripts.run_opex_complete --help
python -m scripts.fill_individual_opex --help
python -m scripts.run_revenue_integration --help
python -m scripts.revenue_preflight --help
python -m scripts.revenue_live_capture_test --help
python -m scripts.inspect_direct_cost_individuals --help
```

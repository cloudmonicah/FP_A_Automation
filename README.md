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

## Run the app

```powershell
python main.py
```

## Project layout

| Path | Contents |
| --- | --- |
| `app.py`, `main.py`, `config.py` | App shell and settings |
| `ui/` | One view per sidebar page |
| `automations/` | Business logic for `forecast`, `actuals` and `budget_setup` |
| `shared/` | Helpers shared across automations |
| `scripts/` | Command-line tools you run by hand (see below) |
| `tests/` | Test scripts |

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

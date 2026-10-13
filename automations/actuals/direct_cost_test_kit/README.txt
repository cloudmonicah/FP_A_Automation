DIRECT COST TEST KIT
1. Back up your project. Extract these files to the FPA_Automation project root.
2. Open PowerShell in the project root. Run: python install_direct_cost_mapping.py
3. Syntax check: python -m py_compile automations\actuals\mapping.py automations\actuals\validation_config.py test_direct_cost.py
4. Run the consolidated test:
 python test_direct_cost.py --aliv "C:\PATH\August 2026.xlsx" --cbl "C:\PATH\1 - CBL Standalone Financials 08.31.26.xlsx" --month 2026-08 --output "C:\Temp\DirectCostTest"
The Tableau file defaults to E:\Workbooks\Direct Cost\Cable Bahamas Group Direct Costs.twb.
Use --open-tableau if you want the runner to launch Tableau first.

IMPORTANT:
- Does NOT change CBL/ALIV financial statement readers or OPEX code.
- Only tests consolidated Tableau versus FS, not individual Tableau extraction.
- Individual mappings are candidate routes disabled pending GL filter validation.
- Category mappings are proposed and must be reconciled against raw Tableau captures.
- If your validation_config.py structure differs from the uploaded version, installer stops without modifying files.
- Existing direct_cost_validation_service.py must already be present.

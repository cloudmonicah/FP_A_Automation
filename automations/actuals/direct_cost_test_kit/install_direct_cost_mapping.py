"""Run from FPA_Automation project root. Safely install Direct Cost mapping without touching OPEX."""
from pathlib import Path
from datetime import datetime
import shutil
import re
root = Path(__file__).resolve().parents[3] if 'automations' in Path(__file__).parts else Path.cwd()
folder = root / 'automations' / 'actuals'
config = folder / 'validation_config.py'
mapping = folder / 'mapping.py'
if not config.exists() or not mapping.exists():
    raise SystemExit('Run this script from the FPA_Automation project root (where automations/actuals exists).')
source = Path(__file__).with_name('direct_cost_mapping.py').read_text(encoding='utf-8')
source = source[source.index('DIRECT_COST_MAPPINGS ='):]
marker = '# DIRECT COST TEST MAPPINGS (managed block)'
end = '# END DIRECT COST TEST MAPPINGS'
original = mapping.read_text(encoding='utf-8')
if marker in original:
    original = original[:original.index(marker)] + original[original.index(end)+len(end):]
else:
    # Remove any previously declared Direct Cost list only when the definition is explicit.
    if re.search(r'^DIRECT_COST_MAPPINGS\s*=', original, re.M):
        raise SystemExit('mapping.py already defines DIRECT_COST_MAPPINGS; merge manually rather than overwrite.')
config_text = config.read_text(encoding='utf-8')
start = config_text.find('DIRECT_COST_MAPPINGS = [')
stop = config_text.find('# REVENUE VALIDATION MAPPINGS', start)
if start < 0 or stop < 0:
    raise SystemExit('Could not identify existing Direct Cost mapping block; no files modified.')
# Replace through the heading before revenue mappings, leaving revenue section intact.
heading_start = config_text.rfind('# ============================================================', start, stop)
# Locate revenue heading by searching backward to nearest full separator preceding heading.
revenue_heading = config_text.rfind('# ============================================================', start, config_text.rfind('# REVENUE VALIDATION MAPPINGS', start, stop))
if revenue_heading < start:
    raise SystemExit('Could not identify Revenue section boundary; no files modified.')
new_config = config_text[:start] + 'from .mapping import DIRECT_COST_MAPPINGS, INDIVIDUAL_DIRECT_COST_MAPPINGS\n\n' + config_text[revenue_heading:]
backup = folder / ('direct_cost_mapping_backup_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
backup.mkdir()
shutil.copy2(mapping, backup / 'mapping.py')
shutil.copy2(config, backup / 'validation_config.py')
mapping.write_text(original.rstrip() + '\n\n' + marker + '\n' + source.rstrip() + '\n' + end + '\n', encoding='utf-8')
config.write_text(new_config, encoding='utf-8')
print('Mapping installed. Backup:', backup)

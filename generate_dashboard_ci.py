"""CI entry point for regenerating index.html on GitHub Actions.

Forces output_html to a path relative to the repo checkout (config.ini's own
paths are absolute Windows paths, meaningless on the runner) and disables the
Excel backup step (there's no local Excel file on the runner at all — that's
fine, read_excel() failing is handled gracefully by run_once() as of
2026-09-14, the Sheet is the primary data source now). Excel comments simply
won't be present in a CI-generated dashboard; everything else is unaffected.
"""
import sys
sys.path.insert(0, '.')
import vaad_bayit_generator as g

_orig_load_config = g.load_config
def _ci_load_config():
    cfg = _orig_load_config()
    cfg.set('paths', 'output_html', 'index.html')
    cfg.set('paths', 'backup_folder', '')
    return cfg
g.load_config = _ci_load_config

g.run_once()

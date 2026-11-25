"""Run the full analysis pipeline from the repository root: python scripts/reproduce_all.py"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STEPS = [
    ("Data integrity checks", "04_verify_data_integrity.py"),
    ("Revision analysis (statistics, sensitivity, supplementary tables)", "05_revision_analysis.py"),
    ("Figures", "01_generate_figures.py"),
    ("Tables", "02_generate_tables.py"),
]

for name, script in STEPS:
    print(f"\n=== {name} ===")
    r = subprocess.run([sys.executable, os.path.join("scripts", script)], cwd=ROOT)
    if r.returncode != 0:
        sys.exit(f"{script} failed with exit code {r.returncode}")
print("\nAll steps completed.")

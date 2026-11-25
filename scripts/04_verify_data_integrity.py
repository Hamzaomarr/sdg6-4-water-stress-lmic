import os
import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_PATH = os.path.join(PROJECT_DIR, 'data', 'AQUASTAT_Project1_Data.xlsx')

def run_audits():
    print("=== Running Project 1 Data Integrity Audits ===")
    df = pd.read_excel(DATA_PATH, sheet_name='Data')
    
    # Check 1: 11 Variables present
    vars_found = df['Variable'].unique().tolist()
    print(f"• Total variables found: {len(vars_found)}")
    assert len(vars_found) == 11, f"Expected 11 variables, found {len(vars_found)}"
    
    # Check 2: Total observations
    print(f"• Total observations: {len(df)}")
    
    # Check 3: Additivity of Water Stress components
    piv = df.pivot(index=['Area', 'Year'], columns='Variable', values='Value').reset_index()
    if 'SDG 6.4.2. Water Stress' in piv.columns:
        ws_tot = piv['SDG 6.4.2. Water Stress']
        ws_sum = (
            piv['SDG 6.4.2. Agricultural Sector Contribution to Water Stress'].fillna(0) +
            piv['SDG 6.4.2. Municipal Sector Contribution to Water Stress'].fillna(0) +
            piv['SDG 6.4.2. Industrial Sector Contribution to Water Stress'].fillna(0)
        )
        diff = np.abs(ws_tot - ws_sum).dropna()
        max_diff = diff.max()
        print(f"• Max Water Stress additivity residual: {max_diff:.6f}")
        assert max_diff < 0.05, f"Water stress additivity failed! Max diff: {max_diff}"
        
    print("✓ All Project 1 data integrity checks passed perfectly!")

if __name__ == '__main__':
    run_audits()

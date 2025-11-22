
import sys
import os
from pathlib import Path
import subprocess

# Add the current directory to sys.path so we can import app
sys.path.append(os.getcwd())

from app import run_optimizer

def test_optimizer():
    print(f"Testing optimizer call with sys.executable: {sys.executable}")
    year_month = "2025年11月"
    
    # Ensure input file exists (it should from previous steps)
    input_file = Path(f"output/{year_month}/{year_month}.xlsx")
    if not input_file.exists():
        print(f"Input file not found: {input_file}")
        return

    print("Calling run_optimizer...")
    result = run_optimizer(year_month)
    print(f"Result: {result}")
    
    if result['success']:
        print("SUCCESS: Optimizer ran successfully.")
    else:
        print(f"FAILURE: Optimizer failed with error: {result.get('error')}")

if __name__ == "__main__":
    test_optimizer()

import os
import subprocess
import sys

def test_dry_run_full_flow(tmp_path):
    """
    Automated Dry-Run Integration Test.
    Creates a temporary environment, runs the script in dry-run,
    and verifies that it completes without errors.
    """
    # 1. Setup minimal environment
    boms_dir = tmp_path / "BOMs"
    boms_dir.mkdir()
    
    # 2. Create sample file
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["System", "Assembly", "Part", "Part_Quantity", "Make o. Buy", "Part_Comments"])
    ws.append(["DT", "Gearbox", "Sun Gear", "1", "m", ""])
    test_file = boms_dir / "test_bom.xlsx"
    wb.save(test_file)
    
    # 3. Prepare env
    os.environ["BOMS_DIR"] = str(boms_dir)
    os.environ["TEAM_ID"] = "999"
    os.environ["DRY_RUN"] = "true"
    os.environ["TEST_MODE"] = "true"
    
    # 4. Run the script in a subprocess
    # We use -m to run as a module if possible, or just call the script
    result = subprocess.run(
        [sys.executable, "main.py", "--system", "DT", "--limit", "1"],
        capture_output=True,
        text=True,
        input="test_bom.xlsx\n\n", # Select file and proceed
    )
    
    assert result.returncode == 0
    assert "Automation finished" in result.stdout

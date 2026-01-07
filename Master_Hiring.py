import subprocess
import sys
import time
from datetime import datetime

# --- CONFIGURATION ---
# The names of your scripts must match exactly what is in your folder
ROBOT_1 = "Robot1.py"  # The Collector (Email -> PDF)
ROBOT_2 = "Robot2.py"  # The Analyst (PDF -> AI -> Excel)
ROBOT_3 = "Robot3.py"  # The Manager (Excel -> Email Offer/Test)
# ---------------------

def log_message(message):
    """Saves progress to a log file so you can prove it works."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    print(log_entry)
    with open("Bot_Activity.log", "a") as log_file:
        log_file.write(log_entry + "\n")

def run_robot(script_name, description):
    """Runs a robot script and waits for it to finish."""
    log_message(f"--- Launching: {description} ---")
    try:
        # This command runs the other python scripts
        result = subprocess.run(
            [sys.executable, script_name],
            capture_output=True, # We capture what the robot prints
            text=True,
            timeout=300 # Stop if it gets stuck for 5 minutes
        )
        
        # If the robot worked (Exit Code 0)
        if result.returncode == 0:
            log_message(f"SUCCESS: {description} finished.")
            # Print the robot's output to the screen just in case
            if result.stdout: print(result.stdout)
            return True
        else:
            log_message(f"ERROR: {description} failed.")
            log_message(f"Error Details: {result.stderr}")
            return False

    except Exception as e:
        log_message(f"CRITICAL ERROR trying to run {script_name}: {e}")
        return False

# --- MAIN PIPELINE ---
def main():
    log_message("=== AITAS MASTER CONTROLLER STARTED ===")

    # Step 1: Download Resumes
    if not run_robot(ROBOT_1, "Robot 1 (Email Ingestion)"):
        log_message("Pipeline STOPPED due to Robot 1 failure.")
        return

    # Step 2: AI Analysis
    if not run_robot(ROBOT_2, "Robot 2 (AI Analysis)"):
        log_message("Pipeline STOPPED due to Robot 2 failure.")
        return

    # Step 3: Make Decisions (Offers/Tests)
    if not run_robot(ROBOT_3, "Robot 3 (Decision Engine)"):
        log_message("Pipeline STOPPED due to Robot 3 failure.")
        return

    log_message("=== PIPELINE COMPLETE. Robot 4 runs manually later. ===")

if __name__ == "__main__":
    main()
"""
Setup helper for Razer Battery Monitor.

Run this once after cloning/downloading:
    python setup_monitor.py

It will:
  1. Install Python dependencies
  2. Optionally create a Windows Startup shortcut so it runs at login
"""

import subprocess
import sys
import os
from pathlib import Path


def install_deps():
    print("Installing dependencies...")
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", "-r",
        str(Path(__file__).parent / "requirements.txt")
    ])
    print("Dependencies installed.\n")


def create_startup_shortcut():
    """Create a .vbs launcher in the Windows Startup folder."""
    startup = Path(os.getenv("APPDATA")) / "Microsoft/Windows/Start Menu/Programs/Startup"
    script_path = Path(__file__).parent.resolve() / "battery_monitor.pyw"
    # Find pythonw.exe — the windowless Python interpreter
    pythonw = Path(sys.executable).parent / "pythonw.exe"
    if not pythonw.exists():
        # Fall back: some installs only have python.exe
        pythonw = Path(sys.executable)

    # Use a .vbs wrapper to launch pythonw silently (no flash of console window)
    vbs_path = startup / "RazerBatteryMonitor.vbs"
    vbs_content = (
        f'Set WshShell = CreateObject("WScript.Shell")\n'
        f'WshShell.Run """{pythonw}"" ""{script_path}""", 0, False\n'
    )

    vbs_path.write_text(vbs_content, encoding="utf-8")
    print(f"Startup shortcut created: {vbs_path}")
    print(f"  Will run: {pythonw} {script_path}\n")


def remove_startup_shortcut():
    startup = Path(os.getenv("APPDATA")) / "Microsoft/Windows/Start Menu/Programs/Startup"
    vbs_path = startup / "RazerBatteryMonitor.vbs"
    if vbs_path.exists():
        vbs_path.unlink()
        print(f"Removed startup shortcut: {vbs_path}")
    else:
        print("No startup shortcut found.")


def main():
    print("=== Razer Battery Monitor Setup ===\n")

    install_deps()

    print("Would you like the monitor to start automatically at login?")
    choice = input("  [Y]es / [N]o / [R]emove existing: ").strip().lower()

    if choice in ("y", "yes"):
        create_startup_shortcut()
    elif choice in ("r", "remove"):
        remove_startup_shortcut()
    else:
        print("Skipped startup shortcut.\n")

    print("Setup complete. To run now:")
    print(f"  pythonw {Path(__file__).parent / 'battery_monitor.pyw'}")
    print()
    print("Options:")
    print("  --threshold 20     # alert at 20% instead of 30%")
    print("  --synapse 3        # force Synapse 3 log path")
    print("  --debug            # enable console logging")


if __name__ == "__main__":
    main()

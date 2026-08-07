# Teams Auto-Typer (Brute Force Edition)

A Python application that uses UI automation (keyboard simulation) to rapidly type student names from an Excel list into Microsoft Teams to automatically add them to a Class Team.

## Download

You can download the compiled Windows executable here:
[Download TeamsAutoTyper.exe](dist/TeamsAutoTyper/TeamsAutoTyper.exe)

## How it works

1. Open the Microsoft Teams desktop app, navigate to your Team, and click **"Add Member"** so the search box is visible on your screen.
2. Open the Auto-Typer (`TeamsAutoTyper.exe`).
3. Select your Excel file containing the class roster.
4. Click the green **START TYPING** button.
5. Quickly switch back to your Microsoft Teams window and click inside the "Add Member" search box.
6. Hands off! The robot takes over, types the names, waits for the directory search, and selects the students automatically.

## Failsafe

If the robot starts typing out of control, **slam your mouse cursor into any corner of your monitor** to trigger an emergency abort.

## Source Code

- `app_brute.py` contains the Tkinter GUI and PyAutoGUI logic.
- `build.ps1` contains the PyInstaller build script.

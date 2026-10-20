@echo off
rem Builds dist\OutlookCalendarReminder.exe (needs Python + pip install pyinstaller pywin32)
python -m PyInstaller --onefile --noconsole --name OutlookCalendarReminder reminder.py

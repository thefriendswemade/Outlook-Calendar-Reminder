# Outlook Calendar Reminder

A small Windows app that pops up a reminder **1 minute before every event** on your Outlook calendar. The popup stays on top of your screen until you dismiss it (or snooze it for 5 minutes), so you don't miss meetings. If it ever can't read Outlook, it shows a red warning instead of failing silently.

<img width="492" height="240" alt="image" src="https://github.com/user-attachments/assets/dc3b7518-502a-4bee-bff7-2013ade255bc" />
<img width="418" height="230" alt="image" src="https://github.com/user-attachments/assets/94392536-bf4c-4e5d-b83b-4023d99bf044" />
<img width="398" height="235" alt="image" src="https://github.com/user-attachments/assets/d07abd9a-51bb-4ff9-8df0-d0ccb78efe12" />

**Requires:** Windows 10/11 and **classic Outlook** (the desktop app with File / Home / Send-Receive tabs). The "new Outlook" and Outlook on the web aren't supported. Outlook needs to be open for reminders to work.

## Install (no Python, no command line)

1. Go to the [Releases page](https://github.com/thefriendswemade/Outlook-Calendar-Reminder/releases/latest) and download **OutlookCalendarReminder.exe**.
2. Put it somewhere permanent (e.g. `Documents`) and double-click it.
3. Click **Yes** when it asks whether to start automatically at login.

That's it. It runs hidden in the background, and it will start itself every time you log in, so you never need to open it again.

> Windows may show a "Windows protected your PC" warning because the app isn't code-signed. Click **More info → Run anyway**.

### Turning it off
Double-click the .exe again and choose **Yes** to turn it off. It will also ask whether to stop it starting at login. Double-click it once more any time to turn it back on.

### Moving or deleting the file
If you move the .exe, turn startup off first (above), then run it from the new location and say Yes again.

## Run it from source (optional)

```
pip install pywin32
python reminder.py          # add --test to show a sample popup
```

### Making the Python version persistent at login

This is how the original was set up, if you'd rather not use the .exe:

1. Install Python and `pip install pywin32`.
2. Press `Win + R`, type `shell:startup`, press Enter.
3. Right-click `run_reminder.vbs` → **Create shortcut**, and move the shortcut into that Startup folder. (The `.vbs` launches `reminder.py` hidden, with no console window.)

To pause it without uninstalling, run `Toggle-Reminder.ps1` (right-click → Run with PowerShell). It creates or removes a `disabled.flag` file and stops or starts the app.

## Building the .exe yourself

```
pip install pyinstaller pywin32
build.bat
```

The result is `dist\OutlookCalendarReminder.exe`. Pushing a tag such as `v1.0.0` builds it automatically on GitHub and attaches it to a Release.

## Troubleshooting

- **No popups:** make sure classic Outlook is running. Errors are logged to `%LOCALAPPDATA%\OutlookCalendarReminder\reminder.log` (the .exe) or `reminder.log` next to the script.
- **Missed a reminder after sleep:** it still alerts for events that started up to 10 minutes ago.

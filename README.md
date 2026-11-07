🔔 Smart Task Notifier

Smart Task Notifier is a lightweight, zero-dependency, 100% native Windows background timer and reminder utility built with VBScript (wscript.exe).

Originally designed to prevent missed office attendance sign-ins and sign-outs, it is fully customizable so you can adapt its timer duration, redirect URL, and popup messages for almost any daily workflow!

✨ Key Features

Zero Dependencies: Runs natively on Windows 7, 10, and 11—no Python, Node.js, or third-party installers required.

Sleep & Reboot Proof: Saves active session timestamps locally (%LOCALAPPDATA%\Smart_Task_Notifier). If your PC goes to sleep, hibernates, or restarts mid-session, the notifier automatically recovers and alerts you right on schedule.

0% CPU Background Tracking: Runs silently in the background without cluttering your taskbar.

Built-in Settings UI: Customize hours, minutes, custom redirect URLs, and popup messages directly from the prompt—no code editing required!

One-Click Web Redirect: Automatically opens your target webpage (HR portal, GitHub, Timesheet, etc.) when starting or completing a session.

🎯 Multi-Purpose Use Cases

While great for attendance regularization, you can configure Smart Task Notifier for many other daily tasks:

Office Attendance & Shift Tracking: Remind yourself to Sign-In on your HR portal (greytHR, Workday, Zoho People, Darwinbox) at login and Sign-Out after 8–9 hours.

Timesheet & Daily Standup Logs: Set a timer to open Jira, Trello, or your daily timesheet tool at the end of your workday.

Pomodoro & Deep Work Sessions: Set a 0 hr 50 min focus timer that opens a break timer or relaxing playlist when finished.

Eye Strain & Hydration Breaks (20-20-20 Rule): Set short intervals (0 hr 30 min or 1 hr 0 min) to step away from your screen and drink water.

Daily GitHub / Coding Check-ins: Keep the default https://github.com URL to remind yourself to push your daily commits or review PRs before shutting down.

Server / Lab / Cloud Instance Shutdowns: Prevent forgotten cloud VMs or test environments from running overnight by setting a reminder to shut them down.

Medication or Stretch Reminders: A reliable desktop popup that stays on top of all windows (System Modal) so you never miss important personal routines.

🚀 How to Install & Run Automatically on Startup

Download SmartTaskNotifier.vbs from this repository.

Press Win + R on your keyboard, type the following path, and press Enter:

shell:startup


(This opens your Windows Startup folder at C:\Users\<YourUsername>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup)

Paste SmartTaskNotifier.vbs directly into that folder.

Double-click SmartTaskNotifier.vbs right away to test it and configure your custom settings! From now on, it will launch automatically every time you turn on or log into your Windows PC.

⚙️ How to Use & Customize

When the popup appears on your screen, you have three options:

[ YES ] — Start / Complete Task:

In the Start Prompt: Opens your configured Redirect URL in your default browser and starts the silent background countdown.

In the Completion Prompt: Opens your Redirect URL so you can log out/submit your work, then resets the timer for tomorrow.

[ NO ] — Close Notifier: Dismisses the reminder and exits cleanly for the day.

[ CANCEL ] — Settings: Opens the interactive configuration menu where you can customize:

Timer Hours (0 to 23)

Timer Minutes (0 to 59)

Redirect URL (Default: https://github.com — change this to your company HR portal, timesheet, or any link)

Start Message (Custom greeting or instruction)

Completion Title (Main alert headline when time is up)

Completion Subtitle (Action reminder text)

Reset All to Default

Note: Your personal settings and custom URLs are saved locally on your computer inside %LOCALAPPDATA%\Smart_Task_Notifier\config.ini, keeping your personal links private and out of the script file.

🛡️ Privacy & Security

100% Local: No data is ever sent to any external server.

No Personal Info Hardcoded: Safe to fork, share, and customize.

Clean Uninstall: Simply delete SmartTaskNotifier.vbs from your Startup folder and remove %LOCALAPPDATA%\Smart_Task_Notifier.

📄 License

Free to use, modify, and distribute under the MIT License.

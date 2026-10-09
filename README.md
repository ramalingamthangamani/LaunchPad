# LaunchPad

One click opens your whole workspace: apps, folders and websites grouped into setups.

## Run
Double-click `LaunchPad.bat` (or `launchpad.pyw`).

## Requirements
Python 3.10+ with: `pip install customtkinter pywin32 winshell`

## Features
- Setups like Work / Study / Gaming, each with its own icon
- Add apps (file picker), folders, files and websites
- Adjustable delay between launches
- "Launch when Windows starts" toggle per setup
- Desktop shortcut per setup that launches it directly
- Everything stored in `profiles.json` next to the app

## Command line
`pythonw launchpad.pyw --launch "Work"` launches a setup without opening the window.

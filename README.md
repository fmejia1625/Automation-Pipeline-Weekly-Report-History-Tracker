# Automation Pipeline Weekly Report History Tracker

This project contains a Python automation script that reads a weekly report configuration and logs historical bot pipeline activity with UTC timestamps.

## What the script does

The script in `bot_history_logger_config_driven_with_utc_v3.py`:
- loads configuration from `bot_history_config_with_utc.json`
- reads weekly pipeline history data
- formats and stores the output in a structured history log
- uses UTC-based timestamps for consistent history tracking
- helps automate weekly reporting for bot pipeline activity

## Files

- `bot_history_logger_config_driven_with_utc_v3.py` - main Python automation script
- `bot_history_config_with_utc.json` - configuration file used by the script
- `Run_Bot_Pipeline_History.bat` - Windows batch launcher for running the script

## How to run

On Windows, you can run the batch file:

```bat
Run_Bot_Pipeline_History.bat
```

Or run the Python script directly:

```powershell
python bot_history_logger_config_driven_with_utc_v3.py
```

## Notes

Make sure the JSON configuration file is present and correctly populated before running the script.

## Current execution setup

The script currently runs on the business user's local machine via Windows Task Scheduler every day at 1:00 PM.

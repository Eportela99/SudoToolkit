# SudoToolkit

Windows tools written in plain PowerShell 5.1 (WinForms GUI). They need no installs and nothing beyond what ships with Windows 10/11.

## Tecnica_System

| Tool | What it does | Run it |
|---|---|---|
| [SMS-Toolkit](Tecnica_System/SMS-Toolkit) | Store maintenance toolkit for SMS store servers and lanes: database maintenance, Mini Ghost, scheduler, scales, remote lane tools, export, Windows repair. Asks for admin rights. | `Run-SMS-Toolkit.bat` |
| [TechRoute](Tecnica_System/TechRoute) | Technician route planner: reads the Helpdesk *Tickets Report* (.xlsx), lists your open tickets, builds a Google Maps route from the office (by open date, by distance or custom) and prints one service form per store. | `Run-TechRoute.bat` |

### Quick start

1. Download the tool's `Run-*.bat` file (open it on GitHub, then **Download raw file**).
2. Double-click it. It downloads the latest version of the script from this repo and starts it, so you always get the newest version.
   If the PC is offline, it starts the last copy it downloaded.

### Build an .exe (optional)

Each tool folder has a `Build-Exe.ps1` (uses [ps2exe](https://github.com/MScholtes/PS2EXE)). Right-click it and choose **Run with PowerShell**. It builds the exe with the SMS-Toolkit icon.

### Notes

- TechRoute looks for the newest `Tickets Report*.xlsx` in its `Assets` folder, its own folder or `Downloads`.
  Ticket reports and ticket PDFs contain customer data and are git-ignored, so never commit them.
- Credits: SMS-Toolkit by Alex C.T & Enrique P.F - Tecnica Systems.

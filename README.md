# KeepRunning
Allows you to launch a program and ensure it stays running. This could be to restart a crashed application or prevent something from being closed accidentally. Useful in a kiosk, library, school, website demo, or web-based application such as a self-registration/signup type of environment.

Use Keep Running as a custom shell so an application such as Internet Explorer is the only available program and automatically re-launch it when closed. Or repurpose old workstations as thin-clients, having them automatically connect your terminal server.

Tested on Windows 2000, XP, Server 2003, Vista, Server 2008, Windows 7

A scripted install/uninstall is not included with this software.

This program runs in the background; without any GUI, taskbar, or system tray icon.

Since this program is 32-bit it can only detect other 32-bit applications.

For a 64-bit version see Keep Running x64
For the opposite of Keep Running see Keep NOT Running

<b>Installation:</b>

1) Ensure this prerequisite is installed:

        Microsoft Visual Basic 6.0 SP6 Run-time Components
        http://jpelectron.com/download/vb6sp6req.exe
        
2) Extract the contents of the .zip file
3) Modify keeprun.ini as indicated below
4) Run keeprun.exe

<b>.ini Settings:</b>

Under most circumstances Detect= and Launch= should be set to the same full path of the executable.

keeprun.exe does not have to directly launch the executable you are detecting, you may specify a batch file or another program which runs the executable being detected.

Do not use quotes around the full path, even if it contains spaces.

Minimally required .ini settings to launch IE on 32-bit Windows

    Detect=C:\Program Files\Internet Explorer\iexplore.exe
    Launch=C:\Program Files\Internet Explorer\iexplore.exe

Minimally required .ini settings to launch IE on 64-bit Windows

    Detect=C:\Program Files (x86)\Internet Explorer\iexplore.exe
    Launch=C:\Program Files (x86)\Internet Explorer\iexplore.exe
      or to launch IE in kiosk mode...
    Launch=C:\Program Files (x86)\Internet Explorer\iexplore.exe -k
      or to launch IE in kiosk mode with a URL...
    Launch=C:\Program Files (x86)\Internet Explorer\iexplore.exe -k http://www.example.com

Minimally required .ini settings to run Iron on 64-bit Windows
   Note: Iron is based on "Chromium" sourcecode, as is Google Chrome

    Detect=C:\Program Files (x86)\SRWare Iron\iron.exe
    Launch=C:\Program Files (x86)\SRWare Iron\iron.exe
      or to launch the browser with no history, in kiosk mode, with a URL...
    Launch=C:\Program Files (x86)\SRWare Iron\iron.exe -incognito -kiosk http://www.example.com

Minimally required .ini settings to launch Terminal Server client on 32-bit Windows

    Detect=C:\Windows\system32\mstsc.exe
    Launch=C:\Windows\system32\mstsc.exe

Minimally required .ini settings to launch Terminal Server client on 64-bit Windows
   Note: You must set the 32-bit client as the default

    Detect=C:\Windows\SysWOW64\mstsc.exe
    Launch=C:\Windows\SysWOW64\mstsc.exe
      or to automatically connect in full-screen mode...
    Launch=C:\Windows\SysWOW64\mstsc.exe /v:[your_terminal_server_ip] /f

<b>Usage:</b>

Optionally, use Autologon to set the workstation to logon automatically.

Set keeprun.exe to start automatically by adding a shortcut in the Startup folder, the registry, or via a Scheduled Task.

Or to have just one program run immediately after login (without explorer, any desktop icons, or a taskbar) set a custom shell...

    For the currently logged on user go to...
       HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System
    For everyone who uses this machine go to...
       HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System

   (If the "System" key does not exist: Edit > New > Key > System)

   Then create a new string value...
   
       Edit > New > String Value
       Value Name: Shell
       Value Data: keeprun.exe (assuming it is located in the Windows working directory)

   When setting a custom shell keeprun.ini should be in the root of the user's profile folder
   usually C:\Documents and Settings\[username]  or  C:\Users\[username]
   Some systems will require that keeprun.ini be located in C:\windows\system32
   In all other cases keeprun.ini should be located in the same folder as keeprun.exe

To log the restarts of a failed application define Launch= as the full path to log-launch.bat and edit this batch file to contain the path to your application.

To restart a service define Detect= as the full path to "your_service.exe" but define Launch= as the full path to service-restart.bat and edit this batch file to contain the Windows service name.

To detect more than one executable on the same system copy keeprun.exe and keeprun.ini into another folder, then use this second instance to detect and launch something else. When running multiple keeprun.exe's sometimes it's best to rename each keeprun.exe differently, like keeprun-for-ie.exe and keeprun-for-other.exe so you can identify each under task manager, the config file should always be named keeprun.ini

# Scripted Win+X shortcuts
Simple scripted backup/restore mechanism for porting your favorite WinX shortcuts to other machines, etc.

## Usage
- faves.md is a markdown format list of your preferred WinX menu entries.
- regenFromFaves.ps1 does exactly that... creates the special WinX menu shortcuts for each entry in your faves.md. MUST BE RUN AS ADMIN.
- list.ps1 dumps out your current Winx "Group3" folder which is the main one typically customized (feel free to change)
  - edit the output into your own faves.md as a way to get started

### handy nuggets:
* [WinX GUI Tool](http://winaero.com/download.php?view.21)
* WinX folder: `cd %LOCALAPPDATA%\Microsoft\Windows\WinX`

### Sample Screenshot
![image](https://cloud.githubusercontent.com/assets/6301228/25764590/7e5a6190-319d-11e7-8724-2fd9222af73f.png)

# Domain Start Menu

Redirected start menus through RoboCopy and inteligent descisions.

## Usage

Before you can use this script you need to configure it by setting the variables and setting up the GPO.

### Variables

_serverStartMenu_ needs to be the UNC path to the root of the start menu on your server.

_localpath_ needs to be the folder on the client computers you want the start menu to be built in.

### GPO Settings

You need to re-direct the users start menu to the folder set in _localpath_, for the best results hide common programs and add them back in through your start menu

## Building the Start Menu

Your Start Menu needs to contain every shortcut you need across all your computers, the script will delete any shortcuts that should point to locally installed software that are not installed on the computer.
All shortcuts pointing to anything other than the C drive will not be checked

Example Start Menu Dir Tree

    \\yourserver\startmenu
    │   Adobe Reader XI.lnk
    │   Intranet.url
    │
    └───Programs
        ├───Accessories
        │       Calculator.lnk
        │       displayswitch.lnk
        │       Math Input Panel.lnk
        │       Paint.lnk
        │       Snipping Tool.lnk
        │       Sound Recorder.lnk
        │       Wordpad.lnk
        │
        ├───Finance
        │       Accounts.lnk
        │
        ├───Office
        │       Word 2010.lnk
        │
        ├───IT
        │       Webmail.lnk
        
## Other Tricks

### 64 Bit Shortcuts

If you have both 64bit and 32 Bit windows on your network you will know the pain of providing shortcuts to software in C:\Program Files (x86), because this solution checks to see if software is installed you can provide both like this:

    Title.lnk -> C:\Program Files (x86)\Title\run.exe
    Title .lnk -> C:\Program Files\Title\run.exe
    
This gives you both shortcuts, one with a space on the end that the user wont be aware of and a normal one. This means that your machines will only have the shortcut for the version they need.

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

D:\USERS\STARTMENU
│   Adobe Reader XI.lnk
│   Espresso Primary.url
│   Internet Explorer.lnk
│   ITP's.url
│   Lancs Outlook Email.url
│   SIMS .net.LNK
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
    ├───Art
    │       Clicker Paint.lnk
    │
    ├───Forgin Languages
    │       Classroom Solutions French.lnk
    │       Pilot Moi Videos.lnk
    │
    ├───Geography
    │       Ace Monkey - EnvironmentKS1.lnk
    │       Ace Monkey - Places.lnk
    │       Barnaby Bear's EAtlas.lnk
    │       Coast 2 Coast.lnk
    │       The River CD.lnk
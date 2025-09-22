<!--
██████╗ ██████╗ ██████╗     ██████╗ ███████╗ ██████╗ ██████╗ ███╗   ██╗███╗   ██╗███████╗ ██████╗████████╗
██╔══██╗██╔══██╗██╔══██╗    ██╔══██╗██╔════╝██╔════╝██╔═══██╗████╗  ██║████╗  ██║██╔════╝██╔════╝╚══██╔══╝
██████╔╝██║  ██║██████╔╝    ██████╔╝█████╗  ██║     ██║   ██║██╔██╗ ██║██╔██╗ ██║█████╗  ██║        ██║   
██╔══██╗██║  ██║██╔═══╝     ██╔══██╗██╔══╝  ██║     ██║   ██║██║╚██╗██║██║╚██╗██║██╔══╝  ██║        ██║   
██║  ██║██████╔╝██║         ██║  ██║███████╗╚██████╗╚██████╔╝██║ ╚████║██║ ╚████║███████╗╚██████╗   ██║   
╚═╝  ╚═╝╚═════╝ ╚═╝         ╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝╚═╝  ╚═══╝╚══════╝ ╚═════╝   ╚═╝   
                                                                                                          
Thank you for using my script. This has been in the background for a while. Finally released it.

You should leave a star in the repo to show your love 💖.

To use the script, you need to install AHK v2.0 from https://www.autohotkey.com/download/ahk-v2.exe

YOU CAN IGNORE EVERYTHING ELSE UNDER THIS, Used for github.
-->
<!-- Info about the macro -->
<div align="center">

# RDP Reconnect
<div align="left">

<!-- Setup tutorial -->

# Setup guide
### 1. Open the script in your choice of text editor.
##### ideally VSCode, but Notepad works fine.
### 2. Configure the "User Config" section of the script.
#### You'll need to crete a folder in your desktop that includes this script, and also the shortcuts that open the RDP sessions. This can be anywhere but ideally you'd have it in Desktop for ease of access
### sample of configuring one user 
#### userConfig.Push({user: "WindowsUsername", shortcut: "filename", title: "RDP window title", id: "", status: ""})
##### user: The username that you'll use for your 
##### shortcut: the shortcut's filename, it's supposed to be a .lnk file, so you'll need to create a shortcut to rdp.exe (or whatever you use)
##### title: This is going to be the RDP window title, this is normally done within the shortcut.
##### id and status: LEAVE THESE BLANK

### General Keybinds
#### F6 - Closes the script
#### F7 - Reloads the script
#### F8 - Testing keybind to see if the macro detects the RDP sessions


<!-- Credits -->
# Credits
#### Creator
 - [Dully176](https://discord.com/users/522940239904243712)

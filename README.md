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
### 1. Create a Folder for the script
- Create a new folder on your desktop, name it whatever you want, I recommend naming it "RDP Sessions"
- Download the script and place it inside the folder you created.
- Inside the same folder, add the shortcuts that open your RDP sesions, they are supposed to be shortcuts that directly launch rdp.exe with some parameters.
#### An example of how it should look like, I am running 6 RDP sessions soooo :D
<img width="924" height="461" alt="image" src="https://github.com/user-attachments/assets/1310bc74-0083-4755-8629-de52669de6b2" />

### 2. Configure the "User Config" section of the script.
- Open the file inside a text editor (VSCode is recommended)
- Go to the "User Config" Section.
- For each RDP session, you'll need to create another "userConfig.push" line and configurate that specifc one to whatever info you have.
```ahk
userConfig.Push({user: "WindowsUsername", shortcut: "filename", title: "RDP window title", id: "", status: ""})
```
#### what each value means
- user: The Windows username for this user's account.
- shortcut: The file name of the shortcut
- title: The title of the RDP window. This is often configured within the shortcut's properties.
- id and status: Leave these blank. The script uses them.
### General Keybinds
- F6: Closes the script.
- F7: Reloads the script.
- F8: A test hotkey to verify that the script is detecting your active RDP sessions.
#### an example of how it should look like based off my screenshot from earlier
```ahk
userConfig.Push({user: "DullyMain", shortcut: "mainAcc", title: "dully natro macro", id: "", status: ""})
userConfig.Push({user: "CharliTadAlt", shortcut: "alt1Acc", title: "charli natro macro", id: "", status: ""})
userConfig.Push({user: "MysticTadAlt", shortcut: "alt2Acc", title: "mystic natro macro", id: "", status: ""})
userConfig.Push({user: "GuidAlt1", shortcut: "GuidAlt1", title: "GuidAlt1", id: "", status: ""})
userConfig.Push({user: "GuidAlt2", shortcut: "GuidAlt2", title: "GuidAlt2", id: "", status: ""})
userConfig.Push({user: "GuidAlt3", shortcut: "GuidAlt3", title: "GuidAlt3", id: "", status: ""})
```
<img width="256" height="244" alt="image (14)" src="https://github.com/user-attachments/assets/9f080d47-12c1-474c-9f79-16e936133824" />
<img width="202" height="177" alt="image" src="https://github.com/user-attachments/assets/8c0623c9-14f8-431b-b42f-e1d7f984a740" />

<!-- Credits -->
# Credits
#### Creator
 - [Dully176](https://discord.com/users/522940239904243712)

# Auto mute speaker(s) for Windows

This app mutes speaker(s) on Windows.


## Features

Once this app is started, it runs as a resident program and shows an icon on the task tray.  
It mutes and/or set volume to zero to the specified speaker(s) at the windows is shutting down.  
Or you can mute and/or set volume to zero to the specified speaker(s) whenever you want.


## Menu

The following menu is shown by clicking the icon in the task tray.

| menu | feature |
|------|---------|
| Settings | Show the setting window. |
| Mute now | Mute speaker(s) and/or set volume to zero now. |
| Exit     | Exit the application and remove the task tray icon. |


## Setting window

When you select 'Settings' menu, the setting window will be shown. You can set the following features.

| Control | Description |
|---------|-------------|
| Mute audio device(s).                 | Mute spcified speaker(s) at the specified timing. |
| Set the volume to zero.               | Set volume of spcified speaker(s) to zero at the specified timing. |
| All speakers are targeted to process. | Mute and/or set volume to zero to all speakers. |
| OK                                    | Save settings and close the window. |
| Cancel                                | Discard settings and close the window.  |

If you don't check 'All speakers are targeted to process', the application mutes and/or set volume to zero to the only current default speaker.


## Process timing

The application mutes and/or set volume to zero at the following timing.

1. When you selct "Mute now" menu.
1. When the windows is shutting down.


## Developing environments

- Windows 11 Pro
- Microsoft Visual C++ Version 17.13.3




# CaptVB
CaptVB is a screenshot utility written in eVB. The user interface is inspired by the [Snipping Tool](https://support.microsoft.com/en-us/windows/use-snipping-tool-to-capture-screenshots-00246869-1843-655f-f220-97299b865f6b) included with modern versions of Windows. If you have ever used Snipping Tool, you will likely feel right at home with CaptVB.

CaptVB has the following features:

- Immediate screen capture or delayed by 3, 5 or 10 seconds.
- Fully configurable save location and file naming pattern including current date and time.
- Optionally ask where to save each screenshot.
- Option to keep the main window always on top.
- Play a sound or flash the notification LED (if available) after screen capture.

![An arrangement of screenshots showing the application in various states. In the first screenshot, the application is shown in the default state. In the second screenshot, the dropdown list containing the screenshot delay values is active. In the third screenshot, the options menu is expanded.](https://github.com/WinCEDev/CaptVB/blob/main/Screenshots/arrangement.png?raw=1)

## Getting Started

When you start the application, the main window is displayed. The main window is comprised of the following components:

- **Screenshot Button:**
Takes the screenshot with the settings you specified.
- **Delay Dropdown:**
Delays the screenshot by 3, 5 or 10 seconds. An icon is displayed in the notification area to indicate the amount of time left. If you want to take the screenshot immediately, just tap on the icon.
- **Options Menu:**
Contains additional settings to make the application best fit your needs.

![The application in default state.](https://github.com/WinCEDev/CaptVB/blob/main/Screenshots/captvb.png?raw=1)

To take a screenshot, simply tap or click on the button with the camera icon.

## The Options Menu

The option menu contains additional settings to make the application best fit your needs.

- **After Capture:**
Contains actions to take after successful capture. 
    - **Play Sound:**
    Plays a sound after the screenshot has been taken, this can be useful if this takes a while on your device.
    - **Flash LED:**
    Flashes the notification LED after the screenshot has been taken. This option may not be visible if your device does not include a notification LED.
- **Always On Top:**
Keep the main window on top of other applications.
- **Ask Before Saving:**
You can select this option to always be asked where you would like to save the screenshot. If this option is not selected, the settings for destination directory and chosen file naming pattern will be used.
- **Destination Directory:**
The location where to save the screenshots.
- **File Naming Pattern:**
The name of the screenshot. You can use the following pattern strings:
    * **{d}** Replaced by the current day.
    * **{m}** Replaced by the current month.
    * **{y}** Replaced by the current year.
    * **{t}** Replaced by the current time (HH:MM).

To revert to the default value, leave the input box blank and click OK, or enter the following value:

```Screenshot {y}-{m}-{d} {t}```

## Advanced: Change App Settings Directly

This section is intended for advanced users and describes how you can change the application settings directly from the registry. Some settings are not available from the user interface because they serve very specific scenarios or are mostly useful for troubleshooting. 

You can find the application settings at the following location:

```HKEY_CURRENT_USER\Software\WinCEDev\CaptVB```

* **FormPos** Position at which the main window appears on application startup. This value contains the x and y coordinates in pixels delimited by a comma.
* **DelayTime** Amount of time before a screenshot is taken. This is not a time value, but rather an index specifying how many seconds to wait. Valid values are:

    - 1: No delay.
    - 2: 3-second delay.
    - 3: 5-second delay.
    - 4: 10-second delay.
* **PlaySound**
Controls whether the option “Play Sound” is enabled.
* **FlashLED**
Controls whether the option “Flash LED” is enabled.
* **AlwaysOnTop**
Controls whether the option “Always on Top” is enabled.
* **AskBeforeSaving**
Controls whether the option “Ask Before Saving” is enabled.
* **SavePath**
Path where to save screenshots.
* **FilePattern**
This contains the file name pattern to use for screenshots to be saved.
* **LEDNum**
Overrides the LED the application will flash on screenshot completion. By default, the first LED (0) will be used. You can change this value if the program is not flashing the expected LED. On some devices, haptics can be activated by specifying the last LED device.
* **SystemSound**
Overrides the [system sound](https://learn.microsoft.com/en-us/windows/win32/multimedia/using-playsound-to-play-system-sounds) to play after the screenshot has been taken. By default, the ‘Asterisk’ sound will be played. If you prefer a different system sound, you can specify one of the following values:
- **SystemAsterisk**
- **SystemDefault**
- **SystemExclamation**
- **SystemExit**
- **SystemHand**
- **SystemQuestion**
- **SystemStart**
- **SystemWelcome**

## Links

- [HPC:Factor Forum Thread](https://www.hpcfactor.com/forums/forums/thread-view.asp?tid=21065&posts=1)
- [HPC:Factor SCL](https://www.hpcfactor.com/scl/2119/WinCEDev/CaptVB/version_0.9.0)

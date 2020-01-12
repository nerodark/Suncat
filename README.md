# Suncat

Suncat is a small Windows service that implements popular cyber-security monitoring concepts. It tracks file system, network, web, UI, keyboard and is flexible enough to be able to track a lot more. All supported events are described below. The source code is provided to be used as a learning tool to understand how the techniques work for IT students, cyber-security enthusiasts and experts. It is currently an entire C# project requiring at least Visual Studio 2019. I take no responsibility for whatever you do with this source code.

# Get started in Visual Studio

As I didn't want to commit the .suo file, you have to setup the solution to use multiple startup projects so you can properly load the application in Debug mode. For that, right-click the solution, select Properties and choose Multiple startup projects option. Make sure that only SuncatServiceTestApp and SuncatHook projects have the Start action and all others are set to None. The projects should always default to the correct action. Click Apply.

# Log customizations

By default, the service, once installed and running on a workstation, logs all activities to the C:\Insights.txt file. You can change that behavior to fit your needs for example if you wish to forward all activities to a SQL database to later pull BI statistics for reports. All the log processing logic is centralized inside the SuncatService\SuncatService.cs class so look at that file and you will catch it pretty easily.

# Events implemented

Event|Data
-|-
CreateFile|Files created on any drive excluding some system folders to reduce noise (supports fixed and removable drives also tags the type of drive on which it happened)
DeleteFile|Files deleted on any drive excluding some system folders to reduce noise (supports fixed and removable drives also tags the type of drive on which it happened)
ChangeFile|Files modified on any drive excluding some system folders to reduce noise (supports fixed and removable drives also tags the type of drive on which it happened)
MoveFile|Files moved on any drive excluding some system folders to reduce noise (supports fixed and removable drives also tags the type of drive on which it happened)
OpenFile|Files opened and read on any drive excluding some system folders to reduce noise (supports fixed, removable drives and network drives also tags the type of drive on which it happened)
CopyFile|Files copied on any drive excluding system some folders to reduce noise (supports fixed and removable drives also tags the type of drive on which it happened)
OpenURL|All URLs visited by popular web browsers (supports Chrome-based, Mozilla-based, Edge browsers even the old Safari for Windows except portable browsers)
SwitchApp|Any window title that became focused and active
TypeOnKeyboard|Concatenates a series of keys until nothing is typed in the last 5 seconds
CopyToClipboard|Text the user copied to your clipboard
HeartBeat|Just tells the system that the service is still running on the tracked computer every 5 minutes
GetPublicIP|Gets the public IP of computer or laptop every 30 minutes using AWS if the IP changed
Explorer|Title of any Explorer window like folder names the user navigated to

# Try, Debug and Test

Select the Debug configuration. Build the solution and start. This will open a console where you will be able to look at the events being tracked live.

# Release and Installer

Download and install the Microsoft Visual Studio Installer Projects extension for Visual Studio to fully build the solution in the Release mode. Just google it or install it through the Extensions menu inside Visual Studio. Then, Select the Release configuration. Build the solution. This will build the MSI file. Then, to create single EXE installer file, navigate to SuncatSetup folder, drag and drop the Release folder on the packager_installer.bat file. Follow the instructions and this will produce the EXE file in the current folder.

A demo installer is available in the releases section of this repository. The demo logs all activities in C:\Insights.txt.

# Limitations

The service will also work on a Windows Server instance but will not track activities of all remote connected users, only locally logged users.
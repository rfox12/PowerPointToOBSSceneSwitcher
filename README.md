# PowerPointToOBSSceneSwitcher
A .NET core based scene switcher that connects to OBS and changes scenes based note meta data. Put "OBS:Your Scene Name" on its own line in your notes and ensure the OBS Web Sockets Server is running and this app will change your scene as you change your PowerPoint slides.

Note this won't build with "dotnet build," instead open a Visual Studio 2019 Developer Command Prompt and build with "msbuild"

This video explains how it works!

[![Watch the video](https://i.imgur.com/v369AtP.png)](https://www.youtube.com/watch?v=ciNcxi2bPwM)

## Usage
* Set a scene for a slide with 
```<language>
OBS:{Scene name as it appears in OBS}
```

Example:
```<language>
OBS:Scenename
```

* Set a default scene (used when a scene is not defined) with
```<language>
OBSDEF:{Scene name as it appears in OBS}
```

Example:
```<language>
OBSDEF:DefaultScene
```

# To Build
1. Go to https://visualstudio.microsoft.com/downloads/?q=build+tools and get the `vs_BuildTools.exe` file.
2. Execute `vs_buildtools.exe --add Microsoft.VisualStudio.Workload.MSBuildTools`
3. A GUI will pop up and you must hit a few buttons.
4. MSBuild.exe will be installed into `C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin` and you need to add that to your PATH
5. Create a new environment variable called `MSBuildSDKsPath` that points to the dotnet SDK on the computer (e.g. `C:\Program Files\dotnet\sdk\5.0.100\Sdks`)
6. You may need to restart VSCode if it has been open during environment variable changes.
7. Open the project in VSCode and in the terminal run `msbuild /t:Restore`
8. Run `msbuild /t:Build /p:Configuration=Release`
 - If you get an error about AxImp.exe missing you need to download the ".NET Framework 4.8 Developer Pack" which has the tools.  https://dotnet.microsoft.com/download/dotnet-framework/thank-you/net48-developer-pack-offline-installer
9. Run `msbuild /t:Publish /p:Configuration=Release`
You should now see the executable under the Release/netcoreapp3.1/publish folder

# To use the program
1. First start both OBS and Powerpoint
2. Then ensure obs-websocket is running on port 4444
3. Simply execute `PowerPointToOBSSceneSwitcher.exe`

The program will crash when you exit a powerpoint presentation



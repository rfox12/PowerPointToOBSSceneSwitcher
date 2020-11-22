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

1. Go to https://visualstudio.microsoft.com/downloads/?q=build+tools and get the `vs_BuildTools.exe` file.
2. Execute `vs_buildtools.exe --add Microsoft.VisualStudio.Workload.MSBuildTools`
3. A GUI will pop up and you must hit a few buttons.
4. MSBuild.exe will be installed into `C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin` and you need to add that to your PATH



# File Collection Utility

Welcome to the File Collection Utility! This tool allows you to seamlessly fetch a set of directories from a Windows machine and download them as a ZIP file.

## Features

- __Robust Implementation:__ Ensures a thorough and recursive retrieval of all files and folders specified as arguments.
- __Fault Tolerance:__ Instead of halting the process, the utility logs any errors to the console, allowing for uninterrupted file retrieval.
- __Shadow Copies:__ To access potentially locked files, our utility creates shadow copies of drives when required. These copies are promptly removed once the file retrieval concludes.

## Live response sessions

Contained in this repository is make_script.ps1. This script serves several functions:

- It embeds the File Collector executable as a base64 string.
- Upon invocation, it unpacks the executable into the current working directory.
- Any command-line arguments passed are directly forwarded to the File Collector executable.

This scripting capability becomes crucial during live response sessions, where direct execution of binaries might be restricted. By transforming the executable into a PowerShell script, this utility ensures that you can still retrieve necessary files even in such constrained environments.

## Compatibility

The resulting script (along with the embedded executable) is compatible with:

- __Windows Versions:__ From Windows 7 SP1 and Windows Server 2008 SP1 onwards.
- __Essential Updates:__
    - __KB3063858:__ Ensures compatibility with the .Net runtime due to new flags in LoadLibraryExW.
    - __KB2999226:__ Incorporates critical CRT changes required by the File Collector executable.
Most Windows versions updated over the past ten years should be compatible. Testing has confirmed functionality on Windows 11 and Windows Server 2008.

## How to build

1. Install the Microsoft .NET SDK 8.0 (Preview) and the Visual Studio build tools
2. Invoke `dotnet publish -c Release` from the root of the repo
3. Run `./make_script` to generate the final script `collect_files.ps1`

## How to run (Example)

`./collect-files "C:\Windows\System32\config" "C:\Users\Administrator" "D:\cool_files" --output files.zip`

## Microsoft Defender Demo
https://github.com/matthieugras/file_collector/assets/35115728/1d80e312-d92e-48fa-b81e-7e8c568ce093


$bytes = [System.IO.File]::ReadAllBytes("$PWD\bin\Release\net8.0\win-x64\publish\FileCollector.exe")
$base64str = [System.Convert]::ToBase64String($bytes)
$template = Get-Content "./script_template.txt";
($template -replace "<EXE_PLACEHOLDER>",$base64str) | Set-Content "collect_files.ps1"
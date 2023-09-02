$b64_section = ""
$exe_bytes = [System.Convert]::FromBase64String($b64_section)
[System.IO.File]::WriteAllBytes("$PWD/program.exe", $exe_bytes)
& "$PWD/program.exe" @args

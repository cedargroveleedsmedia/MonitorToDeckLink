@echo off
echo Inspecting DeckLink interop GUIDs...
echo.
powershell -Command "& { $a = [System.Reflection.Assembly]::LoadFile('%cd%\lib\DeckLinkAPI.dll'); $types = $a.GetTypes() | Where-Object { $_.IsInterface -or $_.IsClass }; foreach ($t in $types) { $attr = $t.GetCustomAttributes([System.Runtime.InteropServices.GuidAttribute], $false); if ($attr.Length -gt 0) { Write-Host $t.Name '=' $attr[0].Value } } }"
echo.
pause

Get-ChildItem "C:\Program Files\Blackmagic Design" -Recurse -Include "*.h","*.idl","*.tlb" -ErrorAction SilentlyContinue | Select-Object FullName
Get-ChildItem "C:\Users\Public\Documents\Blackmagic Design" -Recurse -Include "*.h","*.idl","*.tlb" -ErrorAction SilentlyContinue | Select-Object FullName

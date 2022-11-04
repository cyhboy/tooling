for /d %%X in (*) do "C:\AppFiles\7-Zip\7z.exe" a "%%X.zip" "%%X\" -xr!node_modules >> zip.log

@echo off
setlocal enableextensions
setlocal enabledelayedexpansion

for %%f in (*_cut.*) do (
  set fname=%%f
  set fname=!fname:_cut=!
  ren "%%f" "!fname!"
  )

pause

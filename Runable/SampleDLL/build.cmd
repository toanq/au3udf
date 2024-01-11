set MINGW_PATH=D:\Softs\MinGW\MinGW

set PATH=%PATH%;%MINGW_PATH%\bin
set PATH=%PATH%;%MINGW_PATH%\msys\1.0\bin
set PATH=%PATH%;%MINGW_PATH%\dll

gcc.exe -shared -o main.dll main.c
pause > NUL
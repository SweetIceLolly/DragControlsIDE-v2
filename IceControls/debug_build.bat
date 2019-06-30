windres --include-dir="..\GCC\include\res" "IceControls\IceControls.rc" -O coff -o "IceControls.res"
g++ "IceControls\IceControls.cpp" "IceControls.res" -mwindows -g -o "IceControls.exe"
pause
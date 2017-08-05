@echo off

set PROTO=%1
set PROTOPATH=proto
set SRC_OUT=.\Temp
set PROTO2CS_PATH=.\proto2cs
set PROTO2CPP_PATH=.\proto2cpp
set PROTO2GO_PATH=.\proto2go

echo.
echo ============== Convert %PROTO% =====================

@echo on

mkdir %SRC_OUT%
cd %SRC_OUT%

protoc --proto_path=..\%PROTOPATH% ..\%PROTOPATH%\%PROTO% --csharp_out=.\
protoc --proto_path=..\%PROTOPATH% ..\%PROTOPATH%\%PROTO% --cpp_out=.\
protoc --proto_path=..\%PROTOPATH% ..\%PROTOPATH%\%PROTO% --go_out=.\

cd ..

move /y %SRC_OUT%\*pb.cc %PROTO2CPP_PATH%
move /y %SRC_OUT%\*pb.h %PROTO2CPP_PATH%

move /y %SRC_OUT%\*.cs %PROTO2CS_PATH%

move /y %SRC_OUT%\*.go %PROTO2GO_PATH%

rd /s /q %SRC_OUT%

echo ==================== Complete !!!!=============================
@echo off
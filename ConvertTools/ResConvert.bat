@echo off

set XLS_NAME=%1
::set SHEET_NAME=%2
::set PROTO_NAME=%3


echo.
echo =========Compilation of %XLS_NAME%.xls=========


::---------------------------------------------------
::第一步，将xls经过xls_deploy_tool转成bin和proto
::---------------------------------------------------
set STEP1_XLS2PROTO_PATH=xls2proto

@echo on
cd %STEP1_XLS2PROTO_PATH%

@echo off
echo TRY TO DELETE TEMP FILES:
del *_pb2.py
del *_pb2.pyc
del *.proto
del *.bin
del *.log
del *.txt

@echo on
::python ..\xls2protobuf_v3.py %SHEET_NAME% ..\xls\%XLS_NAME%.xls %PROTO_NAME%
python ..\xls2protobuf_v3.py ..\xls\%XLS_NAME%.xls


::---------------------------------------------------
::第二步：把proto翻译成cs
::---------------------------------------------------
cd ..

set STEP2_PROTO2CS_PATH=.\proto2cs
set PROTO_DESC=proto.protodesc
set SRC_OUT=.\

cd %STEP2_PROTO2CS_PATH%

@echo off
echo TRY TO DELETE TEMP FILES:
del *.cs
del *.protodesc
del *.txt


@echo on
dir ..\%STEP1_XLS2PROTO_PATH%\*.proto /b  > protolist.txt

@echo on
for /f "delims=." %%i in (protolist.txt) do protoc --descriptor_set_out=%PROTO_DESC% --proto_path=..\%STEP1_XLS2PROTO_PATH% ..\%STEP1_XLS2PROTO_PATH%\%%i.proto
::for /f "delims=." %%i in (protolist.txt) do ProtoGen\protogen -i:%PROTO_DESC% -o:%%i.cs
for /f "delims=." %%i in (protolist.txt) do protoc --proto_path=..\%STEP1_XLS2PROTO_PATH% ..\%STEP1_XLS2PROTO_PATH%\%%i.proto --csharp_out=%SRC_OUT%

cd ..

::---------------------------------------------------
::第三步：将bin和cs拷到Assets里
::---------------------------------------------------

@echo off
set OUT_PATH=..\..\..\..\Client\MODWorkspace\MODUnityProject\Assets
set DATA_DEST=StreamingAssets\DataConfig
set CS_DEST=Plugins\ResData


@echo on
copy %STEP1_XLS2PROTO_PATH%\*.bin %OUT_PATH%\%DATA_DEST%
copy %STEP2_PROTO2CS_PATH%\*.cs %OUT_PATH%\%CS_DEST%

::---------------------------------------------------
::第四步：清除中间文件
::---------------------------------------------------
REM @echo off
echo TRY TO DELETE TEMP FILES:
REM cd %STEP1_XLS2PROTO_PATH%
REM del *_pb2.py
REM del *_pb2.pyc
REM del *.proto
REM del *.bin
REM del *.log
REM del *.txt
REM cd ..
REM cd %STEP2_PROTO2CS_PATH%
REM del *.cs
REM del *.protodesc
REM del *.txt
REM cd ..

::---------------------------------------------------
::第五步：结束
::---------------------------------------------------
REM cd ..

@echo on

pause
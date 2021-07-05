call npx vbsnext QRCode.vbs

call del /f test\test*.bmp
call del /f test\test*.svg
call rm -f test/test*.bmp
call rm -f test/test*.svg

call cscript.exe //nologo build/QRCode-bundle.vbs /data:"https://google.ca" /out:"test\test-qr.bmp"

rem tests extracted from Exmple.bat
rem ecr ->  L, M(default), Q, H
rem colordepth -> 1, 24(default)

call CScript.exe build\QRCode-bundle.vbs /data:"Hello World" /out:"test\test-qrcode1.bmp"
call CScript.exe build\QRCode-bundle.vbs /data:"Hello World" /out:"test\test-qrcode2.bmp" /forecolor:#0000FF /backcolor:#E0FFFF /ecr:L /scale:5 /colordepth:1
call CScript.exe build\QRCode-bundle.vbs /data:"Hello World" /out:"test\test-qrcode3.svg"
call CScript.exe build\QRCode-bundle.vbs "test\test.txt" /out:"test\test-qrcode4.bmp"

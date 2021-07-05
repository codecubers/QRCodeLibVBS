call del /f test\units-*.bmp
call rm -f test/units-*.bmp
call npx vbsnext test\\Units.vbs 
call cscript //nologo build\\Units-bundle.vbs
@echo off

set ser=http://10.61.180.168:8501/
set backup_ser=http://10.247.2.61:8501/

echo opening TADA with google chrome, please make sure you have chrome installed.
echo Searching for valid server link, please wait... 

curl %ser%
if ErrorLevel 1 (
	echo Failure
	start chrome %backup_ser%
) else (
	echo Success
	start chrome %ser%
)
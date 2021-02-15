@echo off

set ser=http://10.61.180.168:8501/
set backup_ser=http://10.247.15.68:8501/

echo opening TADA with google chrome, please make sure you have chrome installed.
echo Searching for valid server link, please wait... 
echo #############################################################
echo if TADA server is down, please contact wang.chen@faurecia.com

curl %ser%
if ErrorLevel 1 (
	echo Failure
	start chrome %backup_ser%
) else (
	echo Success
	start chrome %ser%
)
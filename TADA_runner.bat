@echo off

set ser=http://10.247.2.29:8501/
set backup_ser=http://10.247.2.29:8502/

curl %ser%
if ErrorLevel 1 (
	echo Failure
	start "" %backup_ser%
) else (
	echo Success
	start "" %ser%
)

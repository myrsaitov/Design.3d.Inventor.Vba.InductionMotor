@echo off
REM Выполняет команду PUSH для текущего репозитория

REM Устанавливает заголовок окна
title Run "PUSH [%CD%]"

REM Запоминает корневую папку
set ROOT_PATH=%cd%

REM Переходит в рабочую папку
cd ./../__Git.Scripts.Windows/__private

REM Выполняет команду PUSH
call __git_CMD.cmd %ROOT_PATH% "git push -u origin master"

pause

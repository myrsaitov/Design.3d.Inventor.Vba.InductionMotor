@ECHO OFF
:: Выполняет команду PUSH для текущего репозитория

:: Устанавливает заголовок окна
TITLE Run "PUSH [%CD%]"

:: Запоминает корневую папку
SET ROOT_PATH=%CD%

:: Переходит в рабочую папку
CD ./../__Git.Scripts.Windows/__private

:: Выполняет команду PUSH
CALL __git_CMD.cmd --PATH %ROOT_PATH% CALL ./../__Git.Scripts.Windows/__private/__git_check_unpushed_and_push.cmd

PAUSE

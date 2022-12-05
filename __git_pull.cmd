@ECHO OFF
:: Выполняет команду PULL для текущего репозитория

:: Устанавливает заголовок окна
TITLE Run "PULL [%CD%]"

:: Запоминает корневую папку
SET ROOT_PATH=%CD%

:: Переходит в рабочую папку
CD ./../__Git.Scripts.Windows/__private

:: Выполняет команду PULL
CALL __git_CMD.cmd --PATH %ROOT_PATH% git pull

PAUSE

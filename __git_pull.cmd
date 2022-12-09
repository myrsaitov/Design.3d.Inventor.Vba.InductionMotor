@ECHO OFF
:: Выполняет команду PULL для текущего репозитория

:: Устанавливает заголовок окна
TITLE Run "PULL [%CD%]"

:: Выполняет команду PULL
CALL ./../__Git.Scripts.Windows/__private/__git_CMD.cmd --PATH %CD% git pull

PAUSE

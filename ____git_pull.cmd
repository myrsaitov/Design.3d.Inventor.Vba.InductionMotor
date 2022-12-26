@ECHO OFF
:: Выполняет команду PULL для текущего репозитория

:: Заголовок окна - путь и имя файла
TITLE %~0

:: Выполняет команду PULL
CALL ./../__Git.Scripts.Windows/__private/__git_CMD.cmd --PATH %CD% git pull

PAUSE

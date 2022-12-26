@ECHO OFF
:: Выполняет команду PUSH для текущего репозитория

:: Заголовок окна - путь и имя файла
TITLE %~0

:: Выполняет команду PUSH
CALL ./../__Git.Scripts.Windows/__private/__git_CMD.cmd --PATH %CD% CALL ./../__Git.Scripts.Windows/__private/__git_check_unpushed_and_push.cmd

PAUSE

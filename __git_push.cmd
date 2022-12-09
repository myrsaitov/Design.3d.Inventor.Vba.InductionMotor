@ECHO OFF
:: Выполняет команду PUSH для текущего репозитория

:: Устанавливает заголовок окна
TITLE Run "PUSH [%CD%]"

:: Выполняет команду PUSH
CALL ./../__Git.Scripts.Windows/__private/__git_CMD.cmd --PATH %CD% CALL ./../__Git.Scripts.Windows/__private/__git_check_unpushed_and_push.cmd %CD%

PAUSE

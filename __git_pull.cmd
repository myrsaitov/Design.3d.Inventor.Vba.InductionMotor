@echo off

REM Выполняемая команда
set GIT_CMD=git pull

REM Нужно для цветных сообщений
setlocal
call :setESC
cls

REM Печатает путь к репозиторию
echo %ESC%[44mPath:%ESC%[0m 
echo %ESC%[33m%CD%%ESC%[0m

REM Печатает команду, которая будет выполнена
echo %ESC%[44mCommand:%ESC%[0m 
echo %GIT_CMD%

echo %ESC%[44mResult message from Git:%ESC%[0m
%GIT_CMD% && (  REM Если результат выполнения удачный

    REM Вывод сообщения "[ OK ]" зелёным
    setlocal
    call :setESC
    echo;
    echo %ESC%[32m[ OK ]%ESC%[0m

)||(  REM Если результат выполнения неудачный
    
    REM Вывод сообщения "[ Error ]" красным
    setlocal
    call :setESC
    echo;
    echo %ESC%[31m[ Error ]%ESC%[0m

)

echo;

pause

REM Нужно для цветных сообщений
goto:eof
:setESC
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set ESC=%%b
  exit /B 0
)
exit /B 0

@ECHO OFF
SETLOCAL
REM Batch file to forward on unprocessed items. This is designed to be added as a scheduled task to the Book Sorter Induction PCs.
REM Requires Curl: https://curl.se/windows/
REM Download and add binpath to the System PATH ennvironment variable
REM Set the IntelligentReturnSystemManagerPassword in the environment with:
REM setx IntelligentReturnSystemManagerPassword password
REM Note: for Task Scheduler to "see" the new environment variable, you probably need to terminate Taskeng.exe (the one running as the user rather than SYSTEM). Or reboot the system.

REM The following allows you to specify an alternative hostname as the first argument
IF NOT [%1]==[] ( SET "INDUCTION_PC_NAME=%1" ) ELSE ( SET "INDUCTION_PC_NAME=localhost" )
SET _DATESTRING=%DATE:~0,2%/%DATE:~3,2%/%DATE:~6,4%
SET _TIMESTRING=%TIME:~0,2%:%TIME:~3,2%
SET _TIMESTRING=%_TIMESTRING: =0%

IF NOT DEFINED IntelligentReturnSystemManagerPassword ECHO %_DATESTRING% %_TIMESTRING%: IntelligentReturnSystemManagerPassword environmental variable not defined >> 3mStoreAndForward.log && EXIT /B 1

REM Authenticate the user
curl -s -o NUL --cookie-jar booksorterjar.txt --data "password=%IntelligentReturnSystemManagerPassword%" "http://%INDUCTION_PC_NAME%/IntelligentReturn/pages/Index.aspx"

REM Assign number of items to process to variable
FOR /f "tokens=3 delims=><" %%G IN ('curl -L -s --cookie booksorterjar.txt --cookie-jar booksorterjar.txt "http://%INDUCTION_PC_NAME%/IntelligentReturn/pages/StoreAndForward.aspx" ^| find "Store & Forward Items"') DO SET "numOfItemsToProcess=%%G"

REM Record number of items for later checking
ECHO %_DATESTRING% %_TIMESTRING%: %numOfItemsToProcess% items to process >> 3mStoreAndForward.log

REM If number of items to process is zero, exit early
IF [%numOfItemsToProcess%]==[0] GOTO CLEANUP

REM Set the Operation Mode to OUT OF SERVICE, while the store/forwarding process is done
curl -L -s -o NUL --cookie booksorterjar.txt --cookie-jar booksorterjar.txt --data "mode=OUT_OF_SERVICE&submit=Set+Mode" "http://%INDUCTION_PC_NAME%/IntelligentReturn/pages/Support.aspx"

REM Store/forward unprocessed items
curl -L -s -o NUL --cookie booksorterjar.txt --cookie-jar booksorterjar.txt "http://%INDUCTION_PC_NAME%/IntelligentReturn/pages/StoreAndForwardStart.aspx"

REM Wait for the store/forward process to complete - should only take a few seconds but wait two minutes to be sure
TIMEOUT /T 120

REM Set the Operation Mode back to NORMAL
curl -L -s -o NUL --cookie booksorterjar.txt --cookie-jar booksorterjar.txt --data "mode=NORMAL&submit=Set+Mode" "http://%INDUCTION_PC_NAME%/IntelligentReturn/pages/Support.aspx"

:CLEANUP
DEL /Q booksorterjar.txt
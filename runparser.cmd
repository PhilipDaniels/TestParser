@echo off
REM Run the TestParser against everything in an input dir, for all output formats.

SET PARSER=TestParser\bin\debug\TestParser.exe
SET DATADIR=C:\temp\TestParserData

%PARSER% /of:%DATADIR%\TestResults.csv %DATADIR%\*.trx %DATADIR%\*.xml
%PARSER% /of:%DATADIR%\TestResults.json %DATADIR%\*.trx %DATADIR%\*.xml
%PARSER% /of:%DATADIR%\TestResults.kvp %DATADIR%\*.trx %DATADIR%\*.xml
%PARSER% /of:%DATADIR%\TestResults.xlsx /bands:41,100 %DATADIR%\*.trx %DATADIR%\*.xml
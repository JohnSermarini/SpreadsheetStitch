set graphicsDir=graphics
set patternsDir=Patterns
:: set readmeName=SpreadsheetStitch/README.txt

:: Create .exe
pyinstaller SpreadsheetStitch.py --clean --onefile -w --icon=graphics\logo.ico

:: Create disribution folder
if exist SpreadsheetStitch\ rmdir SpreadsheetStitch /s /q
mkdir SpreadsheetStitch
mkdir SpreadsheetStitch\%graphicsDir%
mkdir SpreadsheetStitch\%patternsDir%

:: Populate disribution folder
xcopy %graphicsDir% SpreadsheetStitch\%graphicsDir% /E
xcopy dist\SpreadsheetStitch.exe SpreadsheetStitch

:: Clean remnants of pyinstaller
rmdir dist /s /q
rmdir build /s /q
if exist SpreadsheetStitch.spec del SpreadsheetStitch.spec

:: Create readme file
::echo For help and information, please see: > %readmeName%
::echo https://github.com/JohnSermarini/SpreadsheetStitch >> %readmeName%
copy README.md SpreadsheetStitch :: Copy readme
ren SpreadsheetStitch\README.md README.txt :: Change it to a .txt to improve accessability

:: Add license
:: TODO
set graphicsDir=graphics
set patternsDir=Patterns

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
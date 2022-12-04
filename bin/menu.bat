@echo off

IF "%1" == "" (GOTO HELP)

:OPTION1
IF "%1" == "ie" (GOTO IE)
IF "%1" == "menu" (GOTO MENU)

:HELP
echo menu.bat [OPTION]
echo OPTION:Å´
echo ie
echo menu
GOTO EOF

:IE
explorer /e,/root,%SHORTCUT_BIN_FOLDER%
GOTO EOF

:MENU
explorer /e,/root,%SHORTCUT_BIN_FOLDER%
GOTO EOF


:EOF


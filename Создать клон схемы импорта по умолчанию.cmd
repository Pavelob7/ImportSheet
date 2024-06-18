set clone_schema_name=DefaultSchemeClone
set clone_schema_path=.\ImportSheetConsole\%clone_schema_name%\

xcopy /e ".\ImportSheetConsole\DefaultScheme\" "%clone_schema_path%"

powershell -command "(gc '%clone_schema_path%\CheckHelper.cs') -replace 'namespace ImportSheetConsole.DefaultScheme', 'namespace ImportSheetConsole.%clone_schema_name%' | Out-File '%clone_schema_path%\CheckHelper.cs'"
powershell -command "(gc '%clone_schema_path%\Helper.cs') -replace 'namespace ImportSheetConsole.DefaultScheme', 'namespace ImportSheetConsole.%clone_schema_name%' | Out-File '%clone_schema_path%\Helper.cs'"
powershell -command "(gc '%clone_schema_path%\ImportDefaultScheme.cs') -replace 'namespace ImportSheetConsole.DefaultScheme', 'namespace ImportSheetConsole.%clone_schema_name%' | Out-File '%clone_schema_path%\ImportDefaultScheme.cs'"
powershell -command "(gc '%clone_schema_path%\ImportHelper.cs') -replace 'namespace ImportSheetConsole.DefaultScheme', 'namespace ImportSheetConsole.%clone_schema_name%' | Out-File '%clone_schema_path%\ImportHelper.cs'"

pause
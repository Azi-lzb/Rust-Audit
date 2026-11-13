@echo off
rem Compatibility entry: package the current AuditTools build through the maintained entry.
call "%~dp0package.bat" %*
exit /b %ERRORLEVEL%

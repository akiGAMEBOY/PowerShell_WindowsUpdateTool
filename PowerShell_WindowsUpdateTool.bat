@ECHO OFF
@REM #################################################################################
@REM # �������@�bWindowsUpdateTool�i�N���p�o�b�`�j
@REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
@REM #--------------------------------------------------------------------------------
@REM # �@�@�@�@�b-
@REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  WindowsUpdateTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
SET RETURNCODE=0
@REM PowerShell Core �C���X�g�[���m�F
SET PSFILEPATH="%~dp0\source\powershell\Main.ps1"
WHERE /Q pwsh
IF %ERRORLEVEL% == 0 (
    @REM PowerShell Core �Ŏ��s����ꍇ
    pwsh -NoProfile -ExecutionPolicy Unrestricted -File %PSFILEPATH%
) ELSE (
    @REM PowerShell 5.1  �Ŏ��s����ꍇ
    powershell -NoProfile -ExecutionPolicy Unrestricted -File %PSFILEPATH%
)

SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%

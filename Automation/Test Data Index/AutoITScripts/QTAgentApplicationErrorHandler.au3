#include <GUIConstants.au3>
#Include <GuiButton.au3>
#include <string.au3>
$Var = True
While $Var
    WinWait("QTAgent32.exe - Application Error","")
	$ApplicationErrorDialog =  ControlGetHandle("QTAgent32.exe - Application Error","OK","[CLASS:Button; INSTANCE:1]")
	_GUICtrlButton_Click($ApplicationErrorDialog)
	_GUICtrlButton_Click($ApplicationErrorDialog)
WEnd
#include <GUIConstants.au3>
#Include <GuiButton.au3>
#include <string.au3>
$Var = True
While $Var
    WinWait("QTAgent32.exe","")
	$hnd = ControlGetHandle("QTAgent32.exe","Close the program","[CLASS:Button; INSTANCE:2]")
	_GUICtrlButton_Click($hnd)
	_GUICtrlButton_Click($hnd)
WEnd




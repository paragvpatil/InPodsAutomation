#include <GUIConstants.au3>
#Include <GuiButton.au3>
#include <string.au3>
$Var = True
While $Var
    WinWait("Script Error","")
	MouseClick("left",657, 488)
	Send("{ENTER}")
WEnd
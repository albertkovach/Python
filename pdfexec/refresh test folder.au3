#include <FileConstants.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

$UntouchedTestFiles = "C:\Users\ORB User\Desktop\260005556611"
$TestedFiles = "C:\Users\ORB User\Desktop\10f\260005556611"


DirRemove($TestedFiles, $DIR_REMOVE)
If @error Then
	MsgBox($MB_SYSTEMMODAL, "", "Deleted with errors, but ok")
Endif

DirCreate($TestedFiles)
DirCopy($UntouchedTestFiles, $TestedFiles, $FC_OVERWRITE)

MsgBox($MB_SYSTEMMODAL, "", "Testing folder refreshed !")


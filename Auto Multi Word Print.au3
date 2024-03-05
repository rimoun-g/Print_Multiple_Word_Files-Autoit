; This script is a GUI application for batch printing of Word files.
; It includes several libraries for GUI creation, file handling, and interaction with Microsoft Word.
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <GuiListBox.au3>
#include <Word.au3>
#include <File.au3>
#include <EditConstants.au3>
#Include <WinAPIEx.au3>
#NoTrayIcon

; The GUI includes buttons for selecting files, printing, and other options.
; It also creates a list view to display the selected files and input boxes for the number of copies and delay.
#Region ### START Koda GUI section ### Form=
$mainForm = GUICreate("Multi Word-Files Printing", 700, 400)
$fileListView = GUICtrlCreateList("", 32, 56, 500, 318, $WS_VSCROLL + $ES_AUTOHSCROLL)

$mainLabel = GUICtrlCreateLabel("Multi Word-Files Printing", 200, 16, 350, 30)
GUICtrlSetFont(-1, 12, 700)

$copiesLabel = GUICtrlCreateLabel("Number of Copies", 570, 230, 350, 30)
$delayLabel = GUICtrlCreateLabel("Delay in Seconds", 570, 285, 350, 30)

$chooseFilesButton = GUICtrlCreateButton("Choose Files", 570, 88, 89, 33)
$chooseOneFileButton = GUICtrlCreateButton("Choose one file", 570, 50, 89, 33)
$printButton = GUICtrlCreateButton("Print", 570, 342, 97, 33)

$copiesInput = GUICtrlCreateInput("1", 600, 250, 25, 20)
$delayInput = GUICtrlCreateInput("0", 600, 310, 25, 20)

$contextMenu = GUICtrlCreateContextMenu()
$listContextMenu = GUICtrlCreateContextMenu($fileListView)

$openFileMenuItem = GUICtrlCreateMenuItem("Open selected file", $fileListView)
$openFolderMenuItem = GUICtrlCreateMenuItem("Open file directory", $fileListView)
$deleteSelectedMenuItem = GUICtrlCreateMenuItem("Delete Selected Item", $fileListView)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

; The main loop of the script waits for GUI events and calls the appropriate functions when buttons are clicked or menu items are selected.
While 1
	$guiEvent = GUIGetMsg()
	Switch $guiEvent
		Case $GUI_EVENT_CLOSE
			Exit
		Case $chooseFilesButton
			selectMultipleFiles()
		Case $chooseOneFileButton
			selectOneFile()
		Case $deleteSelectedMenuItem
			deleteSelectedItem()
		Case $printButton
			printFiles()
		Case $openFolderMenuItem
			openFolder()
		Case $openFileMenuItem
			openFile()
	EndSwitch
WEnd

; Functions in the script:
; - deleteSelectedItem(): Deletes the selected item from the list view.
func deleteSelectedItem()
   _GUICtrlListBox_DeleteString($fileListView, _GUICtrlListBox_GetCurSel($fileListView))
EndFunc

; - selectMultipleFiles(): Opens a file dialog to select multiple files and adds them to the list view.
Func selectMultipleFiles()
	$dialog = FileOpenDialog("Select Files", @ScriptDir & "\", "Allowed Files (*.doc;*.docx;*.pdf;*.txt)", 4)
	if $dialog = "" Then
	   Sleep(100)
	else
		$names = StringSplit($dialog, "|")
		$directory = $names[1]
		For $i = 2 to $names[0]
		  _GUICtrlListBox_InsertString($fileListView, $directory & "\" & $names[$i])
		Next
	EndIf
EndFunc

; - selectOneFile(): Opens a file dialog to select one file and adds it to the list view.
func selectOneFile()
	$dialog = FileOpenDialog("Select Files", @ScriptDir & "\", "Allowed Files (*.doc;*.docx;*.pdf;*.txt)")
	if $dialog = "" Then
	   Sleep(100)
	else
		_GUICtrlListBox_InsertString($fileListView, $dialog)
	EndIf
EndFunc

; - printFiles(): Prints the selected files with the specified number of copies and delay.
Func printFiles()
   $delay = guictrlread($delayInput)
   Sleep($delay * 1000)
   $copies = guictrlread($copiesInput)
   $count = _GUICtrlListBox_GetCount($fileListView)

   if $count = 1 Then
	   _FilePrint(_GUICtrlListBox_GetText($fileListView, 0))
	   MsgBox(0, "Done", "Printing Done")
   EndIf

   if $count > 1 Then
	   For $i = 1 to $count
		  _FilePrint(_GUICtrlListBox_GetText($fileListView, $i - 1))
	   Next
	   MsgBox(0, "Done", "Printing Done")
   EndIf

   If $count = 0 Then
	  MsgBox(0, "", "The number of Files should be more than 0")
   EndIf
EndFunc

; - openFolder(): Opens the directory of the selected file.
Func openFolder()
   $count = _GUICtrlListBox_GetCount($fileListView)
   if $count < 1 Then
	  MsgBox(0, "Error", "List is Empty")
   Else
		$selectedFile = _GUICtrlListBox_GetText($fileListView, _GUICtrlListBox_GetCurSel($fileListView))
		_WinAPI_ShellOpenFolderAndSelectItems($selectedFile)
   EndIf
EndFunc

; - openFile(): Opens the selected file.
Func openFile()
   $count = _GUICtrlListBox_GetCount($fileListView)
   if $count < 1 Then
	  MsgBox(0, "Error", "List is Empty")
   Else
		$selectedFile = _GUICtrlListBox_GetText($fileListView, _GUICtrlListBox_GetCurSel($fileListView))
		ShellExecute($selectedFile)
   EndIf
EndFunc

Option Explicit

Dim ObjShortcut, IntResult
Set ObjShortcut = New ClsShortcut
IntResult = ObjShortcut.Run()
Set ObjShortcut = Nothing
Call WScript.Quit(IntResult)

Const CstStrShortcutDir  = "%AppData%\Microsoft\Windows\SendTo"
Const CstStrShortcutExe  = "OpenCsvByExcel.exe"
Const CstStrShortcutName = "Open CSV by Excel"

Class ClsShortcut
	Private ObjFSO'As FileSystemObject
	Private ObjShell'As WshShell
	Private StrShortcutPath
	
	'Constructor
	Private Sub Class_Initialize()
		Set ObjFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set ObjShell = WScript.CreateObject("WScript.Shell")
		StrShortcutPath = ObjFSO.BuildPath(_
			ObjShell.ExpandEnvironmentStrings(CstStrShortcutDir),_
			CstStrShortcutName & ".lnk")
		Call Debug.WriteLine(StrShortcutPath)
	End Sub
	
	'Create or Delete shortcut
	Public Function Run()
		Run = 1
		
		If ObjFSO.FileExists(StrShortcutPath) Then
			If Confirm("Delete confirm", "Are you sure you want to delete a shortcut?") Then
				Call Delete()
				Call MsgBox("Deleted: " & CstStrShortcutName, vbOKOnly + vbInformation, "Delete success")
				Run = 0
			End If
		Else
			If Confirm("Create confirm", "Are you sure you want to create a shortcut?") Then
				Call Create()
				Call MsgBox("Created: " & CstStrShortcutName, vbOKOnly + vbInformation, "Create success")
				Run = 0
			End If
		End If
	End Function
	
	'Create shortcut
	Private Sub Create()
		Dim StrTargetPath
		StrTargetPath = ObjFSO.BuildPath(ObjFSO.GetParentFolderName(WScript.ScriptFullName), CstStrShortcutExe)
		Call Debug.WriteLine(StrTargetPath)
		
		Dim ObjCreator'As IWshShortcut_Class
		Set ObjCreator = ObjShell.CreateShortcut(StrShortcutPath)
		With ObjCreator
			.TargetPath = StrTargetPath
			Call .Save()
		End With
		Set ObjCreator = Nothing
	End Sub
	
	'Delete shortcut
	Private Sub Delete()
		Call ObjFSO.DeleteFile(StrShortcutPath)
	End Sub
	
	'Confirm dialog
	Private Function Confirm(PrmStrTitle, PrmStrMessage)
		Confirm = MsgBox(PrmStrMessage, vbYesNo + vbQuestion, PrmStrTitle) = vbYes
		Call Debug.WriteLine(Confirm)
	End Function
	
	'Destructor
	Private Sub Class_Terminate()
		Set ObjShell = Nothing
		Set ObjFSO = Nothing
	End Sub
End Class

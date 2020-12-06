VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fileFolderHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fso As Scripting.FileSystemObject
Private objFolder As Scripting.folder
Private objFile As Scripting.File
Private dialog As Office.FileDialog
Private Sub Class_Initialize()

    Set fso = New Scripting.FileSystemObject

End Sub
Public Property Get obj_fso() As Scripting.FileSystemObject

    If (fso Is Nothing) Then
        Err.Raise 513, "fileFolderHandler.fso", "No filesystem object"
    Else
        Set obj_fso = fso
    End If

End Property
Public Function existsFile(filePath As String)

    existsFile = fso.FileExists(filePath)

End Function
Public Property Get obj_file() As Scripting.File

    If (objFile Is Nothing) Then
        Err.Raise 513, "fileFolderHandler.objFile", "No file chosen"
    Else
        Set obj_file = objFile
    End If

End Property
Public Function deleteIfFound(fileName As String)

    If (existsFile(fileName)) Then
        fso.DeleteFile fileName
    End If

End Function
Public Function chooseFile() As Boolean

    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With dialog
        .Filters.Clear
        .Title = "Choose a file"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            Set objFile = fso.GetFile(.SelectedItems.Item(1))
        Else
            MsgBox "No file selected"
            chooseFile = False
            Exit Function
        End If
    End With
    
    chooseFile = True

End Function
Public Sub chooseFolder()

    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)

End Sub
Public Function pickedStatus()

    pickedStatus = "file:" & Not IsNull(objFile) & vbCrLf & "folder:" & Not IsNull(objFolder)
    

End Function

Attribute VB_Name = "Module3"
Sub FolderNames_InCurrent()
'Update 20141027
Application.ScreenUpdating = False
Dim xPath As String
Dim xWs As Worksheet
Dim fso As Object, j As Long, folder1 As Object, folder2 As Object
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose the folder"
    .Show
End With
On Error Resume Next

xPathO = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"
MsgBox xPathO
Set xWs = Application.ActiveSheet
xWs.Cells(2, 1).Resize(1, 1).Value = Array("Path")
xWs.Cells(2, 1).Resize(1, 1).Interior.Color = 65535
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder1 = fso.getFolder(xPathO)
For Each folder2 In folder1.SubFolders
    'MsgBox folder2.Name
    For Each SubFolder In folder2.SubFolders
        xRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
        xPath = SubFolder.Path
        Dim arr() As String, arrlen As Integer
        arr = Split(xPath, "\")
        arrlen = UBound(arr) - LBound(arr) + 1
        xWs.Cells(xRow, 1).Resize(1, arrlen).Value = arr
        xWs.Hyperlinks.Add Anchor:=Range(Cells(xRow, arrlen).Address(0, 0)), Address:=xPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder1 = fso.getFolder(xPath)
        getSubFolder folder1
    Next SubFolder
Next folder2
xWs.Cells(2, 1).Resize(1, 1).EntireColumn.AutoFit
Application.ScreenUpdating = True
End Sub

Sub getSubFolder(ByRef prntfld As Object)
Dim SubFolder As Object
Dim subfld As Object
Dim xRow As Long
Set xWs = Application.ActiveSheet
For Each SubFolder In prntfld.SubFolders
    xRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
    xPath = SubFolder.Path
    Dim arr() As String, arrlen As Integer
    arr = Split(xPath, "\")
    arrlen = UBound(arr) - LBound(arr) + 1
    xWs.Cells(xRow, 1).Resize(1, arrlen).Value = arr
    xWs.Hyperlinks.Add Anchor:=Range(Cells(xRow, arrlen).Address(0, 0)), Address:=xPath
Next SubFolder
For Each subfld In prntfld.SubFolders
    getSubFolder subfld
Next subfld
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Utils.PZip"
Option Explicit

Private Type TPZip
    Shell As Object
    Zip As Object
End Type

Private this As TPZip

Private Sub Class_Initialize()
    Set this.Shell = CreateObject("Shell.Application")
End Sub

Public Sub SetZip(ByVal ZipFilePath As String)
    Set this.Zip = GetFolder(ZipFilePath)
    If this.Zip Is Nothing Then
        Me.Create ZipFilePath
        Set this.Zip = GetFolder(ZipFilePath)
    End If
End Sub

Public Sub Unzip(ByVal ToPath As String)
    If this.Zip.Items.Count = 0 Then Exit Sub

    Dim DestinationFolder As Object: Set DestinationFolder = GetFolder(ToPath)
    If DestinationFolder Is Nothing Then
        MkDir ToPath
        Set DestinationFolder = GetFolder(ToPath)
    End If

    DestinationFolder.CopyHere this.Zip.Items, 16
End Sub

Public Sub Create(ByVal Path As String)
    Dim EOCD(22) As Byte
    EOCD(0) = &H50
    EOCD(1) = &H4B
    EOCD(2) = &H5
    EOCD(3) = &H6

    Dim ZipFile As Integer: ZipFile = FreeFile

    Open Path For Binary Access Write As #ZipFile
    Put #ZipFile, , EOCD
    Close #ZipFile
End Sub

Public Sub AddFiles(ParamArray FilesToAdd() As Variant)
    Dim FileToAdd As Variant
    For Each FileToAdd In FilesToAdd
        this.Zip.CopyHere (FileToAdd), 16
    Next
End Sub

Public Sub AddFile(ByVal FileToAdd As String)
    this.Zip.CopyHere (FileToAdd), 16
End Sub

Public Sub AddFolderContent(ByVal FolderToAdd As String)
    Dim Folder As Object: Set Folder = GetFolder(FolderToAdd)
    this.Zip.CopyHere Folder.Items, 16

    Do While this.Zip.Items.Count <> Folder.Items.Count
         Application.Wait (Now + TimeValue("0:00:01"))
    Loop
End Sub

Private Function GetFolder(ByVal Path As String) As Object
    Set GetFolder = this.Shell.Namespace(CStr(Path))
End Function

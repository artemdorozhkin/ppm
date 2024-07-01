VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.Utils.PZip"
Option Explicit

Private Type TPZip
  #If DEV Then
    Shell As Shell32.Shell
    Zip As Shell32.Folder
  #Else
    Shell As Object
    Zip As Object
  #End If
  Path As String
End Type

Private this As TPZip

Private Sub Class_Initialize()
    Set this.Shell = NewShell()
End Sub

Public Property Get Size() As Long
    Size = FileSystem.FileLen(this.Path)
End Property

Public Property Get UnpackedSize() As Long
    Dim Total As Long
    Dim Item As Shell32.FolderItem
    For Each Item In this.Zip.Items
        Total = Total + Item.Size
    Next

    UnpackedSize = Total
End Property

Public Property Get FilesCount() As Long
    FilesCount = this.Zip.Items.Count
End Property

Public Sub SetZip(ByVal ZipFilePath As String)
    this.Path = ZipFilePath
    Set this.Zip = GetFolder(this.Path)
    If this.Zip Is Nothing Then
        Me.Create this.Path
        Set this.Zip = GetFolder(this.Path)
    End If
End Sub

Public Sub Unzip(ByVal ToPath As String)
    If this.Zip.Items.Count = 0 Then Exit Sub

  #If DEV Then
    Dim DestinationFolder As Scripting.Folder: Set DestinationFolder = GetFolder(ToPath)
  #Else
    Dim DestinationFolder As Object: Set DestinationFolder = GetFolder(ToPath)
  #End If
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
  #If DEV Then
    Dim Folder As Shell32.Folder: Set Folder = GetFolder(FolderToAdd)
  #Else
    Dim Folder As Object: Set Folder = GetFolder(FolderToAdd)
  #End If
    this.Zip.CopyHere Folder.Items, 16

    Do While this.Zip.Items.Count <> Folder.Items.Count
         Application.Wait (Now + TimeValue("0:00:01"))
    Loop
End Sub

#If DEV Then
  Private Function GetFolder(ByVal Path As String) As Shell32.Folder
#Else
  Private Function GetFolder(ByVal Path As String) As Object
#End If
    Set GetFolder = this.Shell.Namespace(CStr(Path))
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Module"
Option Explicit

Private Type TModule
    Item As VBComponent
End Type

Private this As TModule

Public Property Get Path() As String
    Path = GetPath()
End Property

Public Property Get Name() As String
    Name = this.Item.Name
End Property

Public Property Get Kind() As ModuleType
    Kind = this.Item.Type
End Property

Public Property Get Item() As VBComponent
    Set Item = this.Item
End Property
Public Property Set Item(ByVal RHS As VBComponent)
    Set this.Item = RHS
End Property

Public Property Get Code() As String
    Code = this.Item.CodeModule.Lines(1, this.Item.CodeModule.CountOfLines)
End Property

Public Sub Rename(ByVal Name As String)
    this.Item.Name = Name
End Sub

Public Sub Move(ByVal ToPath As String)
    Dim FromPath As String: FromPath = GetPath()
    If Strings.Len(FromPath) = 0 Then Exit Sub

    Dim FolderLine As Long: FolderLine = GetFolderAttributeLine()
    If FolderLine = -1 Then Exit Sub

    Dim FolderAttribute As String
    FolderAttribute = this.Item.CodeModule.Lines(FolderLine, 1)
    FolderAttribute = Strings.Replace(FolderAttribute, FromPath, ToPath)
    this.Item.CodeModule.DeleteLines FolderLine, 1
    this.Item.CodeModule.InsertLines FolderLine, FolderAttribute
End Sub

Private Function GetPath() As String
    Dim FolderLine As Long: FolderLine = GetFolderAttributeLine()
    If FolderLine = -1 Then Exit Function
        
    Dim FolderAttribute As String
    FolderAttribute = this.Item.CodeModule.Lines(FolderLine, 1)

    Dim FirstQuote As Long
    FirstQuote = Strings.InStr(1, FolderAttribute, """") + 1
    Dim LastQuote As Long
    LastQuote = Strings.InStrRev(FolderAttribute, """")
    GetPath = Strings.Mid(FolderAttribute, FirstQuote, LastQuote - FirstQuote)
End Function

Private Function GetFolderAttributeLine() As Long
    Dim FolderLine As Long
    If this.Item.CodeModule.Find("'@Folder", FolderLine, 1, this.Item.CodeModule.CountOfLines, 255) Then
        GetFolderAttributeLine = FolderLine
    Else
        GetFolderAttributeLine = -1
    End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UnassignedVariableUsage, ProcedureNotUsed
'@Folder "PearPMProject.src.Module"
Option Explicit

Private Type TModule
  #If DEV Then
    Item As VBComponent
  #Else
    Item As Object
  #End If
End Type

Private this As TModule

Public Property Get Path() As String
    Path = GetPath()
End Property

Public Property Get Name() As String
    Name = this.Item.Name
End Property
Public Property Let Name(ByVal RHS As String)
    this.Item.Name = RHS
End Property

Public Property Get Kind() As ModuleType
    Kind = this.Item.Type
End Property

#If DEV Then
  Public Property Get Item() As VBComponent
#Else
  Public Property Get Item() As Object
#End If
    Set Item = this.Item
End Property
#If DEV Then
  Public Property Set Item(ByVal RHS As VBComponent)
#Else
  Public Property Set Item(ByVal RHS As Object)
#End If
    Set this.Item = RHS
End Property

Public Property Get Code() As String
    If this.Item.CodeModule.CountOfLines = 0 Then Exit Property
    Code = this.Item.CodeModule.Lines(1, this.Item.CodeModule.CountOfLines)
End Property
Public Property Let Code(ByVal RHS As String)
    If this.Item.CodeModule.CountOfLines > 0 Then
        this.Item.CodeModule.DeleteLines 1, this.Item.CodeModule.CountOfLines
    End If
    this.Item.CodeModule.AddFromString RHS
End Property

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

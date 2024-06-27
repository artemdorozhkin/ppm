VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyncCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Sync"
Option Explicit

Implements ICommand

Private Type TSyncCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TSyncCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/sync"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "sync"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim ms As Double: ms = DateTime.Timer

    Dim SyncPaths As Dictionary: Set SyncPaths = NewDictionary()
    Me.CollectPaths SelectedProject.Folder, SyncPaths

    Dim DelNames As Dictionary: Set DelNames = NewDictionary()

    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim Module As Module
    Dim Name As Variant
    For Each Name In SyncPaths
        Dim Path As String: Path = SyncPaths(Name)
        Dim Ext As String: Ext = FSO.GetExtensionName(Path)
        Immediate.WriteLine Name, Ext
        Set Module = NewModule(SelectedProject.GetComponent(Name))
        Dim Code As String: Code = Me.ReadCode(Path)

        If IsFalse(Module) And Not PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            SelectedProject.ImportComponent Path
        ElseIf IsFalse(Module) And PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            ' TODO: need logic for creation the sheet and paste the code
            Information.Err.Raise 5, Information.TypeName(Me), "Non implemented"
        ElseIf IsFalse(Module) Then
            SelectedProject.ImportComponent Path
        ElseIf IsTrue(Module) And PStrings.IsEqual(Ext, "doccls", vbTextCompare) Then
            Module.Code = Code
        Else
            Const MAX_NAME_LEN As Integer = 31
            Dim DelName As String: DelName = Strings.Left("d" & Name, MAX_NAME_LEN)
            DelNames.Add DelName, Empty
            Module.Name = DelName
            SelectedProject.ImportComponent Path
            SelectedProject.RemoveComponent DelName
        End If
    Next

    Dim Component As VBComponent
    Dim Components As VBComponents: Set Components = SelectedProject.Components
    For Each Component In Components
        If DelNames.Exists(Component.Name) Then
            Components.Remove Component
        End If
    Next

    ms = DateTime.Timer - ms
    Immediate.WriteLine PStrings.FString( _
        CLI.Lang.GetValue("messages/synchronized"), _
        SyncPaths.Count, _
        Utils.ConvertTime(ms) _
    )
End Sub

Public Sub CollectPaths(ByVal FolderPath As String, ByRef Paths As Dictionary)
    Dim Folder As Folder: Set Folder = NewFolder(FolderPath)
    Dim SubFolder As Folder
    For Each SubFolder In Folder.SubFolders
        CollectPaths SubFolder.Path, Paths
    Next

    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim File As File
    For Each File In Folder.Files
        Dim Ext As String: Ext = Strings.LCase(FSO.GetExtensionName(File.Path))
        Dim Name As String: Name = FSO.GetBaseName(File.Path)
        If PArray.IncludesAny(Array("bas", "cls", "doccls", "frm"), Ext) Then
            Paths(Name) = File.Path ' Because Name is a unique value, we can't have two identical modules
        End If
    Next
End Sub

Public Function ReadCode(ByVal Path As String) As String
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim Ext As String: Ext = Strings.LCase(FSO.GetExtensionName(Path))
    Dim Name As String: Name = FSO.GetBaseName(Path)
    Dim VBAttr As String

    Select Case Ext
        Case "bas": VBAttr = "Attribute VB_Name = "
        Case Else: VBAttr = "Attribute VB_Exposed = "
    End Select

    Dim Code As String: Code = PFileSystem.ReadFile(Path)
    Dim AttrPosition As String: AttrPosition = Strings.InStr(1, Code, VBAttr, vbTextCompare)
    Code = Strings.Right(Code, Strings.Len(Code) - AttrPosition)

    Dim NewLinePosition As String: NewLinePosition = Strings.InStr(1, Code, vbNewLine)
    Code = Strings.Right(Code, Strings.Len(Code) - NewLinePosition - 1)

    ReadCode = Code
End Function

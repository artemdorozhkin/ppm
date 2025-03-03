VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Export"
Option Explicit

Implements ICommand

Private Type TExportCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/export"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "export"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("encoding") = CLI.Lang.GetValue("param", "name=encoding")
    this.CommandInfo.Params.Item("save-struct") = CLI.Lang.GetValue("param", "name=save-struct")
    this.CommandInfo.Params.Item("path") = CLI.Lang.GetValue("param", "name=path")
    this.CommandInfo.Params.Item("no-clear") = CLI.Lang.GetValue("param", "name=no-clear")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If Not SelectedProject.IsSaved Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/saveRequired")
        End
    End If

    Dim ms As Double: ms = DateTime.Timer
    Dim Encoding As String
    Encoding = Utils.GetTokenOrConfigValue("encoding", this.Tokens)

    Dim SaveStruct As Boolean
    SaveStruct = Utils.GetTokenOrConfigValue("save-struct", this.Tokens, CastTo:=VbVarType.vbBoolean)

    Dim Folder As String
    Folder = SelectedProject.Folder

    Dim UserFolder As String
    UserFolder = Utils.GetTokenOrConfigValue("path", this.Tokens)
    If Strings.Len(UserFolder) > 0 Then
        Folder = UserFolder
    End If

    Dim NoClear As Boolean
    NoClear = Utils.GetTokenOrConfigValue("no-clear", this.Tokens, CastTo:=VbVarType.vbBoolean)

    Dim Exporter As ProjectExporter
    Set Exporter = NewProjectExporter( _
        Destination:=Folder, _
        SaveStruct:=SaveStruct, _
        RewriteLastExport:=Not NoClear _
    )
    Exporter.ExportAll Encoding

    ms = DateTime.Timer - ms
    Immediate.WriteLine PStrings.FString( _
        CLI.Lang.GetValue("messages/exported"), _
        Exporter.Components.Count, _
        IIf(Exporter.Components.Count > 1, "s", ""), _
        ConvertTime(ms), _
        SelectedProject.Folder _
    )
End Sub

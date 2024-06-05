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
    Config As Config
    Tokens As Tokens
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "export"
    this.CommandInfo.Description = "Exports the selected project."
    this.CommandInfo.Params.Item("encoding") = "\\t\\tExport files with specified encoding."
    this.CommandInfo.Params.Item("save-struct") = "\\tSave the RubberDuck structure when exporting a project."
    this.CommandInfo.Params.Item("path") = "\\t\\t\\tFolder to export."
    this.CommandInfo.Params.Item("no-clear") = "\\t\\t\\tDo not delete files from the last export."
    this.CommandInfo.Usage = Array("[<options>]")
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
    Dim Encoding As String: Encoding = Me.GetEncoding()

  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
  #End If
    Dim SaveStruct As Boolean
    SaveStruct = this.Config.GetValue("save-struct")

    Dim Folder As String
    Folder = SelectedProject.Folder
    Dim UserFolder As String
    UserFolder = this.Config.GetValue("path")
    If Strings.Len(UserFolder) > 0 Then
        Folder = UserFolder
    End If

    Dim NoClear As Boolean: NoClear = this.Config.GetValue("no-clear")

    Dim Exporter As ProjectExporter
    Set Exporter = NewProjectExporter( _
        Destination:=Folder, _
        SaveStruct:=SaveStruct, _
        RewriteLastExport:=Not NoClear _
    )
    Exporter.ExportAll Encoding

    ms = DateTime.Timer - ms
    Immediate.WriteLine PStrings.FString( _
        "exported {0} module{1} in {2}\\nroot folder: {3}\\n", _
        Exporter.Modules.Count, _
        IIf(Exporter.Modules.Count > 1, "s", ""), _
        ConvertTime(ms), _
        SelectedProject.Folder _
    )
End Sub

Public Function GetEncoding() As Variant
    Dim EncodingDef As Definition: Set EncodingDef = Definitions("encoding")
    If this.Tokens.IncludeDefinition(EncodingDef) Then
        GetEncoding = this.Tokens.GetTokenByDefinition(EncodingDef).Text
    Else
        GetEncoding = EncodingDef.Default
    End If
End Function

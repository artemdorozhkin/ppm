VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Init"
Option Explicit

Implements ICommand

Private Type TInitCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TInitCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "init"
    this.CommandInfo.Description = "Create a package module."
    this.CommandInfo.Params.Item("name") = "\\t\\tSet project name."
    this.CommandInfo.Params.Item("yes") = "\\tSkips dialog and sets default values."
    this.CommandInfo.Usage = Array("<package name>")
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
    Dim AfterDialog As Boolean: AfterDialog = this.Config.GetValue("_after-dialog")
    Dim SkipByDefault As Boolean: SkipByDefault = this.Config.GetValue("yes")

    If Not AfterDialog Then InitDialog.SetDefault
    If this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        InitDialog.SetNameAndDefault this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text
    ElseIf Not SkipByDefault Then
        InitDialog.SetDefault
    End If

    If SkipByDefault Then
        Dim ms As Double: ms = DateTime.Timer
        Me.CreatePack
        ms = DateTime.Timer - ms
        Immediate.WriteLine PStrings.FString("package is initialized in {0}", ConvertTime(ms))
    Else
        InitDialog.Start
    End If
End Sub

Public Sub CreatePack()
    If SelectedProject.IsModuleExists("package") Then
        Immediate.WriteLine "module already exists"
        Exit Sub
    End If

  #If DEV Then
    Dim PackModule As VBComponent: Set PackModule = SelectedProject.AddModule("package")
  #Else
    Dim PackModule As Object: Set PackModule = SelectedProject.AddModule("package")
  #End If
    With PackModule.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString PStrings.FString( _
            "' {\\n" & _
            "'   ""name"": ""{0}"",\\n" & _
            "'   ""version"": ""{1}"",\\n" & _
            "'   ""description"": ""{2}"",\\n" & _
            "'   ""author"": ""{3}"",\\n" & _
            "'   ""git"": ""{4}""\\n" & _
            "' }", _
            PackInfo.Name, PackInfo.Version, PackInfo.Description, PackInfo.Author, PackInfo.GitURL _
        )
    End With
End Sub

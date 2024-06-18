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
    Lang As Lang
End Type

Private this As TInitCommand

Private Sub Class_Initialize()
    Set this.Lang = NewLang(Russian)
    this.Lang.SetBlock "/root/commands/init"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "init"
    this.CommandInfo.Description = this.Lang.GetValue("description")
    this.CommandInfo.Params.Item("name") = this.Lang.GetValue("param", "name=name")
    this.CommandInfo.Params.Item("yes") = this.Lang.GetValue("param", "name=yes")
    this.CommandInfo.Usage = Array(this.Lang.GetValue("usage"))
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
        Immediate.WriteLine PStrings.FString(this.Lang.GetValue("initialized"), ConvertTime(ms))
    Else
        InitDialog.Start
    End If
End Sub

Public Sub CreatePack()
    If SelectedProject.IsModuleExists("package") Then
        Immediate.WriteLine this.Lang.GetValue("exists")
        Exit Sub
    End If

  #If DEV Then
    Dim PackModule As VBComponent: Set PackModule = SelectedProject.AddModule("package")
  #Else
    Dim PackModule As Object: Set PackModule = SelectedProject.AddModule("package")
  #End If
    Dim JB As PJSONBuilder: Set JB = New PJSONBuilder
    JB.AddObject _
        "name", PackInfo.Name, _
        "version", PackInfo.Version, _
        "description", PackInfo.Description, _
        "author", PackInfo.Author, _
        "git", PackInfo.GitURL

    With PackModule.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Utils.CommentString(JB.ToString())
        .InsertLines 1, PStrings.FString("'@Folder(""{0}"")", SelectedProject.Name)
    End With
End Sub

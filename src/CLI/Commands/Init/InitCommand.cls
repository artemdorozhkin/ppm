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
    CLI.Lang.SetBlock "/root/ppmCommands/init"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "init"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("name") = CLI.Lang.GetValue("param", "name=name")
    this.CommandInfo.Params.Item("yes") = CLI.Lang.GetValue("param", "name=yes")
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
    Dim AfterDialog As Boolean
    AfterDialog = GetFirstValueFrom("_after-dialog", this.Tokens, this.Config)
    Dim SkipByDefault As Boolean
    SkipByDefault = GetFirstValueFrom("yes", this.Tokens, this.Config)

    Dim Pack As Pack
    If SelectedProject.IsComponentExists("package") Then
        Set Pack = NewPack(SelectedProject.GetComponent("package"))
    End If

    Dim HaveProjectName As Boolean
    HaveProjectName = this.Tokens.IncludeTokenKind(TokenKind.Identifier)
    Dim InitialName As String
    If HaveProjectName Then
        InitialName = this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text
    End If

    If IsTrue(Pack) And Not AfterDialog Then
        If Not HaveProjectName Then
            InitialName = Pack.Name
        End If
        InitDialog.SetDefault _
            Name:=InitialName, _
            Version:=Pack.Version, _
            Description:=Pack.Description, _
            Author:=Pack.Author, _
            GitURL:=Pack.Git
    ElseIf Not AfterDialog Then
        InitDialog.SetDefault
    ElseIf HaveProjectName Then
        InitDialog.SetNameAndDefault InitialName
    ElseIf Not SkipByDefault Then
        InitDialog.SetDefault
    End If

    If SkipByDefault Then
        Dim ms As Double: ms = DateTime.Timer
        Me.CreatePack
        ms = DateTime.Timer - ms
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/initialized"), ConvertTime(ms))
    Else
        InitDialog.Start
    End If
End Sub

Public Sub CreatePack()
  #If DEV Then
    Dim PackComponent As VBComponent
  #Else
    Dim PackComponent As Object
  #End If
    If SelectedProject.IsComponentExists("package") Then
        Set PackComponent = SelectedProject.GetComponent("package")
    Else
        Set PackComponent = SelectedProject.AddComponent("package")
    End If

    Dim JB As PJSONBuilder: Set JB = New PJSONBuilder
    JB.AddObject _
        "name", PackInfo.Name, _
        "version", PackInfo.Version, _
        "description", PackInfo.Description, _
        "author", PackInfo.Author, _
        "git", PackInfo.GitURL, _
        "references", Me.GetReferences()

    SelectedProject.Rename PackInfo.Name & "Project"

    With PackComponent.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Utils.CommentString(JB.ToString())
        .InsertLines 1, PStrings.FString("'@Folder(""{0}"")", SelectedProject.Name)
    End With
End Sub

Public Function GetReferences() As Object
    Dim References As Object
    Set References = NewDictionary()

    Dim r As Object
    For Each r In ThisWorkbook.VBProject.References
        References(r.Name) = FString("{0}.{1}", r.Major, r.Minor)
    Next

    Set GetReferences = References
End Function

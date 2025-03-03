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

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim AfterDialog As Boolean
    AfterDialog = this.Tokens.IncludeDefinition(Definitions("_after-dialog"))
    Dim SkipByDefault As Boolean
    SkipByDefault = this.Tokens.IncludeDefinition(Definitions("yes"))

    Dim Pack As Pack
    If SelectedProject.IsComponentExists("package") Then
        Set Pack = NewPack(SelectedProject.GetComponent("package"))
    End If

    Dim HaveProjectName As Boolean
    HaveProjectName = this.Tokens.IncludeTokenKind(TokenKind.Identifier)
    Dim InitialName As String
    If HaveProjectName Then
        InitialName = this.Tokens.PopFirstTokenKind(TokenKind.Identifier).Text
    End If

    If IsTrue(Pack) And Not AfterDialog Then
        InitPack = InitDialog.SetDefault( _
            Name:=GetFirstTrue(InitialName, Pack.Name), _
            Version:=Pack.Version, _
            Description:=Pack.Description, _
            Author:=Pack.Author, _
            GitURL:=Pack.Git _
        )
    ElseIf HaveProjectName Then
        InitPack = InitDialog.SetDefault(Name:=InitialName)
    Else
        InitPack = InitDialog.SetDefault(Name:=SelectedProject.Name)
    End If

    If Not AfterDialog And SkipByDefault Then
        ResultPack = InitPack
    End If

    If SkipByDefault Then
        Dim ms As Double: ms = DateTime.Timer
        Dim PackComponent As Object
        Set PackComponent = Me.CreatePackComponent()

        If IsTrue(Pack) Then
            Dim CurrentPack As Pack
            Set CurrentPack = NewPack(PackComponent)
            CurrentPack.Update Dependencies:=Pack.Dependencies, DevDependencies:=Pack.DevDependencies
        End If

        ResultPack = InitDialog.SetDefault( _
            Name:="", _
            Version:="", _
            Description:="", _
            Author:="", _
            GitURL:="" _
        )
        ms = DateTime.Timer - ms
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/initialized"), ConvertTime(ms))
    Else
        InitDialog.Start
    End If
End Sub

Public Function CreatePackComponent() As Object
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
        "name", ResultPack.Name, _
        "version", ResultPack.Version, _
        "description", ResultPack.Description, _
        "author", ResultPack.Author, _
        "git", ResultPack.GitURL, _
        "references", Me.GetReferences()

    SelectedProject.Rename ResultPack.Name & "Project"

    With PackComponent.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Utils.CommentString(JB.ToString())
        .InsertLines 1, PStrings.FString("'@Folder(""{0}"")", SelectedProject.Name)
    End With

    Set CreatePackComponent = PackComponent
End Function

Public Function GetReferences() As Object
    Dim References As Object
    Set References = NewDictionary()

    Dim r As Object
    For Each r In SelectedProject.References
        References(r.Name) = FString("{0}.{1}", r.Major, r.Minor)
    Next

    Set GetReferences = References
End Function

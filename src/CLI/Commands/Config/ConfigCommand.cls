VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Config"
Option Explicit

Implements ICommand

Private Type TConfigCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
    SelectedScope As ConfigScopes
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/config"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "config"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("location") = CLI.Lang.GetValue("param", "name=location")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim ProjectConfigPath As String
    ProjectConfigPath = ConfigPaths.GetProjectConfigPath()
    Dim FSO As Object
    Set FSO = NewFileSystemObject()
    Dim HasProjectConfig As Boolean
    HasProjectConfig = FSO.FileExists(ProjectConfigPath)
    If HasProjectConfig Then
        this.SelectedScope = ConfigScopes.ProjectScope
    Else
        this.SelectedScope = ConfigScopes.UserScope
    End If

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.PopFirstTokenKind(TokenKind.SubCommand)

    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    Else
        ppm "config --help"
    End If

    If this.Tokens.IncludeDefinition(Definitions("location")) Then
        this.SelectedScope = Me.GetScopeByLocation()
        If this.SelectedScope = ConfigScopes.ProjectScope And _
        Not HasProjectConfig Then
            PFileSystem.SaveToFile ProjectConfigPath, ""
        End If
    End If
    Config.ReadScope this.SelectedScope

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If

    Select Case SubCommandName
        Case "get"
            ConfigGet

        Case "set"
            ConfigSet

        Case "delete"
            ConfigDelete

        Case "edit"
            ConfigEdit

        Case "list"
            Config.ReadFileByScope this.SelectedScope
            ConfigList

        Case "generate"
            If Not SelectedProject.IsSaved Then
                CLI.Lang.GetValue ("messages/needSaveProject")
                Exit Sub
            End If

            Config.GenerateDefault ConfigScopes.ProjectScope
            Immediate.WriteLine PStrings.FString( _
                CLI.Lang.GetValue("messages/fileGenerated"), ConfigPaths.GetProjectConfigPath() _
            )

    End Select
End Sub

Public Function GetScopeByLocation() As ConfigScopes
    Dim Location As String
    Location = this.Tokens.PopTokenByDefinition(Definitions("location")).Text

    Dim Scope As ConfigScopes
    Select Case Location
        Case "project"
            Scope = ConfigScopes.ProjectScope

        Case "user"
            Scope = ConfigScopes.UserScope

        Case "global"
            Scope = ConfigScopes.GlobalScode

        Case Else
            Information.Err.Raise _
                9, _
                Information.TypeName(Me), _
                PStrings.FString(CLI.Lang.GetValue("messages/unknownConfigType"), Location)

    End Select

    GetScopeByLocation = Scope
End Function

Public Sub ConfigSet()
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        If KeyToken.Kind <> TokenKind.Identifier Then GoTo Continue
        i = i + 1
        If i > this.Tokens.Count Then Exit For
        
        Dim ValueToken As SyntaxToken
        Set ValueToken = this.Tokens.MatchToken( _
            i, _
            TokenKind.Identifier, _
            PStrings.FString(CLI.Lang.GetValue("messages/expectedValue"), KeyToken.Text) _
        )
        Config.SetValue KeyToken.Text, ValueToken.Text
        Immediate.WriteLine PStrings.FString("{0}={1}", KeyToken.Text, ValueToken.Text)
Continue:
    Next

    Config.Save this.SelectedScope
End Sub

Public Sub ConfigGet()
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        Dim Key As String: Key = KeyToken.Text
        If KeyToken.Kind = TokenKind.Identifier Then
            Immediate.WriteLine PStrings.FString("{0}={1}", Key, Config.GetValue(Key))
        End If
    Next
End Sub

Public Sub ConfigDelete()
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        If KeyToken.Kind = TokenKind.Identifier Then
            Config.DeleteKey KeyToken.Text
            Config.Save this.SelectedScope
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/deleted"), KeyToken.Text)
            Exit Sub
        End If
    Next

    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/unknownKey"), KeyToken.Text)
End Sub

Public Sub ConfigEdit()
    If Config.IsMissing(this.SelectedScope) Then
        Config.GenerateDefault this.SelectedScope
    End If
    Interaction.Shell "notepad.exe " & ConfigPaths.GetByScope(this.SelectedScope), vbNormalFocus
End Sub

Public Sub ConfigList()
    Immediate.WriteLine Config.GetText()
End Sub

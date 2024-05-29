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
    Config As Config
    Tokens As Tokens
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "config"
    this.CommandInfo.Description = "Manage the ppm configuration file."
    this.CommandInfo.Params.Item("global") = "\\tUse global config."
    this.CommandInfo.Params.Item("location") = "\\tUse config from specify location <global|user|project>."
    this.CommandInfo.Usage = Array( _
        "set <key>=<value>", _
        "get <key>=<value>", _
        "edit", _
        "delete <key>" _
    )
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
    Dim Config As ConfigIO: Set Config = GetConfig()

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.GetFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    Else
        ppm "config --help"
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If

    Select Case SubCommandName
        Case "get"
            ConfigGet Config

        Case "set"
            ConfigSet Config

        Case "delete"
            ConfigDelete Config

        Case "edit"
            ConfigEdit Config

        Case "list"
            ConfigList Config

    End Select
End Sub

Public Function GetConfig() As ConfigIO
    If this.Config.GetValue("global") Then
        Set GetConfig = Configs.GetGlobalConfig()
        Exit Function
    End If

    Dim Location As String: Location = this.Config.GetValue("location")
    If IsFalse(Location) Then
        Set GetConfig = Configs.GetProjectConfig()
    Else
        Select Case Location
            Case "project"
                Set GetConfig = Configs.GetProjectConfig()

            Case "user"
                Set GetConfig = Configs.GetUserConfig()

            Case "global"
                Set GetConfig = Configs.GetGlobalConfig()

            Case Else
                Information.Err.Raise 9, Information.TypeName(Me), "Unknown config type: " & Location
        End Select
    End If
End Function

Public Sub ConfigSet(ByVal Config As ConfigIO)
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        If KeyToken.Kind <> TokenKind.Identifier Then GoTo Continue
        i = i + 1
        If i > this.Tokens.Count Then Exit For
        
        Dim ValueToken As SyntaxToken
        Set ValueToken = this.Tokens.MatchToken( _
            i, TokenKind.Identifier, "Expected Value token after Key token " & KeyToken.Text _
        )
        Config.CreateOrWrite KeyToken.Text, ValueToken.Text
        Immediate.WriteLine PStrings.FString("{0}={1}", KeyToken.Text, ValueToken.Text)
Continue:
    Next
End Sub

Public Sub ConfigGet(ByVal Config As ConfigIO)
  #If DEV Then
    Dim Data As Dictionary: Set Data = Config.Read()
  #Else
    Dim Data As Object: Set Data = Config.Read()
  #End If

    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        Dim Key As String: Key = KeyToken.Text
        If KeyToken.Kind = TokenKind.Identifier Then
            Immediate.WriteLine PStrings.FString("{0}={1}", Key, Data(Key))
        End If
    Next
End Sub

Public Sub ConfigDelete(ByVal Config As ConfigIO)
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        If KeyToken.Kind = TokenKind.Identifier Then
            Config.DeleteKey KeyToken.Text
            Immediate.WriteLine PStrings.FString("key '{0}' deleted", KeyToken.Text)
        End If
    Next
End Sub

Public Sub ConfigEdit(ByVal Config As ConfigIO)
    Config.AddDataWithoutReplace Definitions.Items
    Interaction.Shell "notepad.exe " & Config.ConfigPath, vbNormalFocus
End Sub

Public Sub ConfigList(ByVal Config As ConfigIO)
    Immediate.WriteLine Config.Text
End Sub

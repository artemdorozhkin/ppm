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
    CLI.Lang.SetBlock "/root/ppmCommands/config"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "config"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("location") = CLI.Lang.GetValue("param", "name=location")
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
    Dim Config As Config: Set Config = GetConfig()

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.PopFirstTokenKind(TokenKind.SubCommand)

    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    Else
        ppm "config --help"
    End If

    If SubCommandName = "generate" Then
        If Not SelectedProject.IsSaved Then
            CLI.Lang.GetValue ("messages/needSaveProject")
            Exit Sub
        End If

        GenerateDeafultConfig
        Immediate.WriteLine PStrings.FString( _
            CLI.Lang.GetValue("messages/fileGenerated"), PFileSystem.BuildPath(SelectedProject.Folder, ".ppmrc") _
        )
        Exit Sub
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

Public Sub GenerateDeafultConfig()
    Dim Config As Config
    Set Config = NewConfig(ConfigScopes.ProjectScope)

    Dim SkipableKeys As Variant
    SkipableKeys = Array( _
        "author-name", _
        "author-url", _
        "dev-dep", _
        "help", _
        "location", _
        "name", _
        "path", _
        "version", _
        "yes" _
    )

    Dim Definition As Variant
    For Each Definition In Definitions.Items.Items
        If PArray.IncludesAny(SkipableKeys, Definition.Key) Then GoTo Continue
        If PStrings.StartsWith(Definition.Key, "_") Then GoTo Continue

        Dim Value As String
        If Information.IsMissing(Definition.Default) Then
            Value = ""
        Else
            Value = Conversion.CStr(Definition.Default)
        End If
        Config.File.CreateOrWrite Definition.Key, Value
Continue:
    Next
End Sub

Public Function GetConfig() As Config
    Dim LocationDef As Definition
    Set LocationDef = Definitions("location")

    Dim Location As String
    If this.Tokens.IncludeDefinition(LocationDef) Then
        Location = this.Tokens.PopTokenByDefinition(LocationDef).Text
    Else
        Set GetConfig = NewConfig(ConfigScopes.ProjectScope)
        Exit Function
    End If

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

    Set GetConfig = NewConfig(Scope)
End Function

Public Sub ConfigSet(ByVal Config As Config)
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
        Config.File.CreateOrWrite KeyToken.Text, ValueToken.Text
        Immediate.WriteLine PStrings.FString("{0}={1}", KeyToken.Text, ValueToken.Text)
Continue:
    Next
End Sub

Public Sub ConfigGet(ByVal Config As Config)
  #If DEV Then
    Dim Data As Dictionary: Set Data = Config.File.Read()
  #Else
    Dim Data As Object: Set Data = Config.File.Read()
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

Public Sub ConfigDelete(ByVal Config As Config)
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        If KeyToken.Kind = TokenKind.Identifier Then
            Config.File.DeleteKey KeyToken.Text
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/deleted"), KeyToken.Text)
        End If
    Next
End Sub

Public Sub ConfigEdit(ByVal Config As Config)
    Config.File.AddDataWithoutReplace Definitions.Items
    Interaction.Shell "notepad.exe " & Config.File.ConfigPath, vbNormalFocus
End Sub

Public Sub ConfigList(ByVal Config As Config)
    Immediate.WriteLine Config.File.Text
End Sub

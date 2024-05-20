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
    Tokens As Tokens
    CommandInfo As CommandInfo
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "config"
    this.CommandInfo.Description = "Manage the ppm configuration file."
    this.CommandInfo.Usage = Array( _
        "set <key>=<value>", _
        "get <key>=<value>", _
        "delete <key>" _
    )
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim Config As ConfigIO: Set Config = Configs.GetGlobalConfig()

    Select Case this.Tokens.GetFirstTokenKind(TokenKind.SubCommand).Text
        Case "get"
            ConfigGet Config

        Case "set"
            ConfigSet Config

        Case "delete"
            ConfigDelete Config

    End Select
End Sub

Private Sub ConfigSet(ByVal Config As ConfigIO)
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        Dim Section As String: Section = "Global"
        Dim Key As String: Key = KeyToken.Text
        If PStrings.Contains(Key, ".") Then
            Dim Parts As Variant: Parts = Strings.Split(Key, ".")
            Section = Parts(0)
            Key = Parts(1)
        End If
        If KeyToken.Kind = TokenKind.Identifier Then
            i = i + 1
            Dim ValueToken As SyntaxToken
            Set ValueToken = this.Tokens.MatchToken( _
                             i, TokenKind.Identifier, "Expected Value token after Key token" _
                                                     )
            Config.WriteData Section, Key, ValueToken.Text
        End If
    Next
End Sub

Private Sub ConfigGet(ByVal Config As ConfigIO)
    Dim GlobalSection As Dictionary
    Set GlobalSection = Config.ReadSection("Global")

    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        Dim Key As String: Key = KeyToken.Text
        Dim Section As Dictionary: Set Section = GlobalSection
        If PStrings.Contains(Key, ".") Then
            Dim Parts As Variant: Parts = Strings.Split(Key, ".")
            Dim CustomSection As String: CustomSection = Parts(0)
            Key = Parts(1)
            Set Section = Config.ReadSection(CustomSection)
        End If
        If KeyToken.Kind = TokenKind.Identifier Then
            Immediate.WriteLine FString("{0}={1}", Key, Section(Key))
        End If
    Next
End Sub

Private Sub ConfigDelete(ByVal Config As ConfigIO)
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim KeyToken As SyntaxToken: Set KeyToken = this.Tokens(i)
        Dim Key As String: Key = KeyToken.Text
        Dim Section As String: Section = "Global"
        If PStrings.Contains(Key, ".") Then
            Dim Parts As Variant: Parts = Strings.Split(Key, ".")
            Section = Parts(0)
            Key = Parts(1)
        End If
        Config.DeleteKey Section, Key
    Next
End Sub

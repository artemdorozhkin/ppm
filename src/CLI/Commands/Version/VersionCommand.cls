VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Version"

Implements ICommand

Private Type TVersionCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TVersionCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/version"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "version"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent("package"))
    Dim NewVers As String

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.PopFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    Else
        Dim NewVersToken As SyntaxToken
        Set NewVersToken = this.Tokens.PopFirstTokenKind(TokenKind.Identifier)
        If IsFalse(NewVersToken) Then
            Immediate.WriteLine PStrings.FString("v{0}", Pack.Version)
            Exit Sub
        End If

        NewVers = NewVersToken.Text
        Dim Parts As Variant: Parts = Strings.Split(NewVers, ".")
        If UBound(Parts) <> 2 Then
            Immediate.WriteLine CLI.Lang.GetValue("messages/incorrectSyntax")
            End
        End If
        Pack.Update Version:=NewVers
        Immediate.WriteLine PStrings.FString("v{0}", NewVers)
        Exit Sub
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If

    Dim CurrentVersParts As Variant: CurrentVersParts = Strings.Split(Pack.Version, ".")

    Select Case SubCommandName
        Case "major"
            CurrentVersParts(0) = CurrentVersParts(0) + 1
            CurrentVersParts(1) = 0
            CurrentVersParts(2) = 0

        Case "minor"
            CurrentVersParts(1) = CurrentVersParts(1) + 1
            CurrentVersParts(2) = 0

        Case "patch"
            CurrentVersParts(2) = CurrentVersParts(2) + 1

        Case Else
            ppm "help version"
            End

    End Select

    NewVers = Strings.Join(CurrentVersParts, ".")
    Pack.Update Version:=NewVers
    Immediate.WriteLine PStrings.FString("v{0}", NewVers)
End Sub

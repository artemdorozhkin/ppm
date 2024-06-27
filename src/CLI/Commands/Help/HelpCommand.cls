VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Help"
Option Explicit

Implements ICommand

Private Type THelpCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
    Config As Config
End Type

Private this As THelpCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/help"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "help"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage/item")
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
    If IsFalse(this.Tokens) Then
        ShowPPMInfo
        Exit Sub
    End If

    If this.Tokens.Count = 0 Then
        ShowPPMInfo
        Exit Sub
    End If

    If this.Tokens.Count = 1 Then
        If IsEqual(this.Tokens(1).Text, this.CommandInfo.Name) Then
            ppm "help -h"
            Exit Sub
        End If
    End If

    On Error GoTo Catch
    Dim CommandName As String: CommandName = Me.GetCommandName()
    If Strings.Len(CommandName) = 0 Then Information.Err.Raise 0

    Dim Command As ICommand
    Set Command = Application.Run(PStrings.FString("New{0}Command", CommandName), this.Config, this.Tokens)
    Immediate.WriteLine Command.CommandInfo.GetUsage()
Exit Sub

Catch:
    CLI.Lang.SetBlock "/root/errors"
    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/unknownCommand"), CommandName)
End Sub

Public Function GetCommandName() As String
    Dim Commands As Variant: Commands = this.Tokens.GetTokensKind(TokenKind.Command)
    If UBound(Commands) = -1 Then
        GetCommandName = this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text
        Exit Function
    End If

    Dim HasHelpOption As Boolean: HasHelpOption = this.Config.GetValue("help")
    Dim CommandIndex As Long
    CommandIndex = IIf(HasHelpOption, 0, 1)

    Dim Command As SyntaxToken: Set Command = Commands(CommandIndex)
    Dim Name As String: Name = CLI.FindCommand(Command.Text)

    If Strings.Len(Name) = 0 Then
        GetCommandName = Command.Text
    Else
        GetCommandName = Name
    End If
End Function

Public Sub ShowPPMInfo()
    CLI.Lang.SetBlock "/root"

    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent("package"))
    Immediate.WriteLine PStrings.FString("Pear Package Manager, v{0}\\n", Pack.Version)
    Immediate.WriteLine CLI.Lang.GetValue("usage")
    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("usageExample"))
    Immediate.WriteLine CLI.Lang.GetValue("commands")

    Dim List As Collection: Set List = New Collection
    Dim Command As Variant
    For Each Command In CLI.Commands
        Dim Line As String
        If Strings.Len(Line) = 0 Then
            Line = Command
        Else
            Line = PStrings.FString("{0}, {1}", Line, Command)
        End If

        If Strings.Len(Line) > 45 Then
            List.Add Line
            Line = ""
        End If
    Next
    If Strings.Len(Line) > 0 Then List.Add Line

    Immediate.WriteLine Strings.Join(PCollection.ToArray(List), vbNewLine)
End Sub

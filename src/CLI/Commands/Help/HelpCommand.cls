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
    Tokens As Tokens
    CommandInfo As CommandInfo
End Type

Private this As THelpCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "help"
    this.CommandInfo.Description = "Shows information about the command."
    this.CommandInfo.Usage = Array("[<command>]")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Public Function GetCommandName() As String
    Dim Commands As Variant: Commands = this.Tokens.GetTokensKind(TokenKind.Command)
    If UBound(Commands) = -1 Then
        GetCommandName = this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text
        Exit Function
    End If

    Dim HelpOptionIndex As Long: HelpOptionIndex = this.Tokens.GetTokenIndex("h", ShortOptionItem)
    If HelpOptionIndex = -1 Then HelpOptionIndex = this.Tokens.GetTokenIndex("help", OptionItem)
    Dim CommandIndex As Long
    CommandIndex = IIf(HelpOptionIndex = -1, 1, 0)

    Dim Command As SyntaxToken: Set Command = Commands(CommandIndex)
    Dim Name As String: Name = CLI.FindCommand(Command.Text)

    If Strings.Len(Name) = 0 Then
        GetCommandName = Command.Text
    Else
        GetCommandName = Name
    End If
End Function

Public Function ShowPPMInfo()
    Immediate.WriteLine FString("Pear Package Manager, v{0}\\n", App.Version)
    Immediate.WriteLine "Usage:"
    Immediate.WriteLine FString( _
        "ppm ""install""\\nppm ""install <package>""\\nppm ""init""\\nppm ""init <package name>""\\n" _
    )
    Immediate.WriteLine "Commands:"

    Dim List As Collection: Set List = New Collection
    Dim Command As Variant
    For Each Command In CLI.Commands
        Dim Line As String
        If Strings.Len(Line) = 0 Then
            Line = Command
        Else
            Line = FString("{0}, {1}", Line, Command)
        End If

        If Strings.Len(Line) > 45 Then
            List.Add Line
            Line = ""
        End If
    Next
    If Strings.Len(Line) > 0 Then List.Add Line

    Immediate.WriteLine Strings.Join(PCollection.ToArray(List), vbNewLine)
End Function

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
    If Strings.Len(CommandName) = 0 Then Err.Raise 0

    Dim Command As ICommand
    Set Command = Application.Run(FString("New{0}Command", CommandName), this.Tokens)
    Immediate.WriteLine Command.CommandInfo.GetUsage()
Exit Sub

Catch:
    Immediate.WriteLine "Unknown command " & CommandName
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Commands.Help"
Option Explicit

Implements ICommand

Private Type THelpCommand
    Args As Variant
    CommandInfo As CommandInfo
End Type

Private this As THelpCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "help"
    this.CommandInfo.Description = "Shows information about the command."
    this.CommandInfo.Usage = Array("[<command>]")
End Sub

Public Property Let Args(ByVal RHS As Variant)
    this.Args = RHS
End Property

Public Function GetCommandName() As String
    If Information.IsEmpty(this.Args) Then Exit Function

    Dim Arg As Variant
    For Each Arg In this.Args
        If IsEqual(Arg, "help", vbTextCompare) Or _
           IsEqual(Arg, "-h", vbTextCompare) Or _
           IsEqual(Arg, "--help", vbTextCompare) Then GoTo Continue
        Exit For
Continue:
    Next

    GetCommandName = Arg
End Function

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim Arg As String: Arg = CLI.FindCommand(Me.GetCommandName())
    If Strings.Len(Arg) = 0 Then Arg = "help"

    On Error GoTo Catch
    Dim Command As ICommand: Set Command = Application.Run(FString("New{0}Command", Arg), this.Args)
    Immediate.WriteLine Command.CommandInfo.GetUsage()
Exit Sub

Catch:
    Immediate.WriteLine "Unknown command " & Arg
End Sub


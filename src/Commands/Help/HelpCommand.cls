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

Public Sub PrintHelp(ByRef Info As CommandInfo)
    Debug.Print FString(Info.Description & "\\n")
    Debug.Print FString("Usage:\\nppm ""{0} {1}""\\n", Info.Name, Strings.Join(Info.Usage, "\\n  "))

    If Info.Params.Count > 0 Then
        Debug.Print "Options:"
        Dim Param As Variant
        For Each Param In Info.Params
            Dim Description As String: Description = Info.Params(Param)
            Debug.Print FString("--{0}{1}", Param, Description)
        Next
    End If

    Dim CommandAliases As Collection: Set CommandAliases = New Collection
    Dim CLIAliases As Object: Set CLIAliases = CLI.Aliases
    Dim Alias As Variant
    For Each Alias In CLIAliases
        If IsEqual(CLIAliases(Alias), Info.Name) Then
            CommandAliases.Add Alias
        End If
    Next

    If CommandAliases.Count > 0 Then
        Debug.Print FString("Aliases:\\n{0}", Strings.Join(ToArray(CommandAliases), ", "))
    End If
End Sub

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

Private Sub ICommand_Run()
    Dim Arg As String: Arg = CLI.FindCommand(Me.GetCommandName())
    If Strings.Len(Arg) = 0 Then Arg = "help"

    On Error GoTo Catch
    Dim Command As ICommand: Set Command = Application.Run(FString("New{0}Command", Arg), this.Args)
    On Error GoTo 0

    PrintHelp Command.CommandInfo
Exit Sub

Catch:
    Debug.Print "Unknown command " & Arg
End Sub

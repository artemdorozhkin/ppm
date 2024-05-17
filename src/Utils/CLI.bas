Attribute VB_Name = "CLI"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Property Get Commands() As Variant
    Commands = Array( _
        "export", _
        "help", _
        "install" _
    )
End Property

Public Property Get Aliases() As Object
    Dim Buffer As Object: Set Buffer = NewDictionary(vbTextCompare)

    Buffer("i") = "install"
    Buffer("add") = "install"

    Buffer("exp") = "export"
    Buffer("save") = "export"

    Set Aliases = Buffer
End Property

Public Function Parse(ByRef Args As Variant) As ICommand
    If Not IsArray(Args) Then
        Set Parse = NewHelpCommand(Empty)
        Exit Function
    ElseIf ArrayIncludes(Args, "-h") Or ArrayIncludes(Args, "--help") Then
        Set Parse = NewHelpCommand(Args)
        Exit Function
    End If

    Dim Command As String: Command = CLI.FindCommand(Args(0))
    If Strings.Len(Command) = 0 Then
        Debug.Print "Unknown command " & Args(0)
        End
    End If

    Set Parse = Application.Run(FString("New{0}Command", Command), Args)
End Function

Public Function FindCommand(ByVal Name As String) As String
    If ArrayIncludes(Commands, Name) Then
        FindCommand = Name
        Exit Function
    End If

    If Aliases.Exists(Name) Then
        FindCommand = Aliases(Name)
        Exit Function
    End If
End Function

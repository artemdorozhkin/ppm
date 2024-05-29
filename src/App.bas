Attribute VB_Name = "App"
'@Folder "PearPMProject.src"
Option Explicit

Public Const Version = "0.2.0"
Public SelectedProject As Project

Public Sub ppm(Optional ByVal StringArgs As String)
    On Error GoTo Catch
    Set SelectedProject = NewProject(Application.VBE.ActiveVBProject)
    Dim Command As ICommand: Set Command = CLI.ParseCommand(NewLexer(StringArgs).Lex())
    Command.Exec
Exit Sub

Catch:
    Immediate.WriteLine PStrings.FString( _
        "[{0}] ERR: #{1} {2}", _
        Information.Err.Source, _
        Information.Err.Number, _
        Information.Err.Description _
    )
End Sub



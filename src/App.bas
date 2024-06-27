Attribute VB_Name = "App"
'@Folder "PearPMProject.src"
Option Explicit

Private Type TApp
    SelectedProject As Project
End Type

Private this As TApp

Public Property Get SelectedProject() As Project
    Set SelectedProject = this.SelectedProject
End Property

Public Sub ppm(Optional ByVal StringArgs As String)
    On Error GoTo Catch
    Set this.SelectedProject = NewProject(Application.VBE.ActiveVBProject)
    CLI.InitLang
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

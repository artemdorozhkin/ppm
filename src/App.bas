Attribute VB_Name = "App"
'@Folder "PearPMProject.src"
Option Explicit

Public Const Version = "0.1.0"
Public SelectedProject As Project

Public Sub ppm(Optional ByVal StringArgs As String)
    On Error GoTo Catch
    Set SelectedProject = NewProject(Application.VBE.ActiveVBProject)
    Dim Args As Variant: Args = Immediate.ParseArgs(StringArgs)
    Dim Command As ICommand: Set Command = CLI.ParseCommand(Args)
    Command.Exec
Exit Sub

Catch:
    Immediate.WriteLine FString("[{0}] ERR: #{1} {2}", Err.Source, Err.Number, Err.Description)
End Sub

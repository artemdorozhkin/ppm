Attribute VB_Name = "App"
'@Folder "PearPMProject.src"
Option Explicit

Public Const Version = "0.1.0"
Public SelectedProject As Project

Public Sub ppm(Optional ByVal StringArgs As String)
    On Error GoTo Catch
    Set SelectedProject = NewProject(Application.VBE.ActiveVBProject)
    Dim Args As Variant: Args = ArgsParser.Parse(StringArgs)
    Dim Command As ICommand: Set Command = CLI.Parse(Args)
    Command.Run
Exit Sub

Catch:
    Debug.Print FString("[{0}] ERR: #{1} {2}", Err.Source, Err.Number, Err.Description)
End Sub

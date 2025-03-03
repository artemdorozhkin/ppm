Attribute VB_Name = "App"
'@Folder "PearPMProject.src"
Option Explicit

Private Type TApp
    SelectedProject As Project
    ThisProject As Project
End Type

Private this As TApp

Public Property Get ThisProject() As Project
    Set ThisProject = this.ThisProject
End Property

Public Property Get SelectedProject() As Project
    Set SelectedProject = this.SelectedProject
End Property

Public Sub ppm(Optional ByVal StringArgs As String)
    On Error GoTo Catch
    Set this.SelectedProject = NewProject(Application.VBE.ActiveVBProject)

    If Config.IsMissing(ConfigScopes.GlobalScode) Then
        Config.GenerateDefault ConfigScopes.GlobalScode
    End If

    If Config.IsMissing(ConfigScopes.UserScope) Then
        Config.GenerateDefault ConfigScopes.UserScope
    End If

    Config.ReadScope

    Dim Project As Project
  #If DEV Then
    Dim VBProject As VBProject
  #Else
    Dim VBProject As Object
  #End If
    For Each VBProject In Application.VBE.VBProjects
        If PStrings.IsEqual(VBProject.Name, "PearPMProject") Then
            Set this.ThisProject = NewProject(VBProject)
            Exit For
        End If
    Next

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

Attribute VB_Name = "ConfigPaths"
'@Folder("PearPMProject.src.Config")
Option Explicit

Public Function GetByScope(ByVal Scope As ConfigScopes) As String
    Select Case Scope
        Case ConfigScopes.GlobalScode
            GetByScope = ConfigPaths.GetGlobalConfigPath()
        Case ConfigScopes.UserScope
            GetByScope = ConfigPaths.GetUserConfigPath()
        Case ConfigScopes.ProjectScope
            GetByScope = ConfigPaths.GetProjectConfigPath()
        Case Else
            Exit Function
    End Select
End Function

Public Function GetProjectConfigPath() As String
    If Not SelectedProject.IsSaved Then Exit Function

    GetProjectConfigPath = PFileSystem.BuildPath(SelectedProject.Folder, ".ppmrc")
End Function

Public Function GetUserConfigPath() As String
    Dim Folder As String
    Folder = PFileSystem.BuildPath(Interaction.Environ("APPDATA"), "ppm")
    On Error Resume Next
    FileSystem.MkDir Folder

    GetUserConfigPath = PFileSystem.BuildPath(Folder, ".ppmrc")
End Function

Public Function GetGlobalConfigPath() As String
    Dim Folder As String
    Folder = PFileSystem.BuildPath(Interaction.Environ("PROGRAMDATA"), "ppm")
    On Error Resume Next
    FileSystem.MkDir Folder

    GetGlobalConfigPath = PFileSystem.BuildPath(Folder, ".ppmrc")
End Function

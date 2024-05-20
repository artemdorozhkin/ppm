Attribute VB_Name = "PPMPaths"
'@Folder("PearPMProject.src.Utils")
Option Explicit

Public Function GetGlobalConfigPath() As String
    With NewFileSystemObject()
        Dim PPMPath As String
        PPMPath = .BuildPath(Interaction.Environ("APPDATA"), "ppm")
        If Not .FolderExists(PPMPath) Then CreateFoldersRecoursive PPMPath
        Dim ConfigPath As String
        ConfigPath = .BuildPath(PPMPath, "config.cfg")
        If Not .FileExists(ConfigPath) Then .CreateTextFile(ConfigPath).Close
        GetGlobalConfigPath = ConfigPath
    End With
End Function

Public Function GetCurrentProjectConfigPath() As String
    On Error GoTo Catch
    Dim ProjectPath As String: ProjectPath = PPMPaths.GetProjectPath()
    On Error GoTo 0
    With NewFileSystemObject()
        Dim ProjectConfig As String
        ProjectConfig = .BuildPath(ProjectPath, "config.cfg")
    End With

    GetCurrentProjectConfigPath = ProjectConfig
Exit Function

Catch:
    If Err.Number = 76 Then
        Immediate.WriteLine "ERR: It is required to save the project before exporting."
        End
    Else
        Err.Raise Err.Number, TypeName("Configs")
    End If
End Function

Public Function GetPPMProjectsPath() As String
    With NewFileSystemObject()
        Dim PPMPath As String
        PPMPath = .BuildPath(Interaction.Environ("LOCALAPPDATA"), "ppm")
        Dim PPMProjectsPath As String
        PPMProjectsPath = .BuildPath(PPMPath, "projects")
        If Not .FolderExists(PPMProjectsPath) Then CreateFoldersRecoursive PPMProjectsPath
        GetPPMProjectsPath = PPMProjectsPath
    End With
End Function

Public Function GetProjectPath() As String
    Dim PPMProjectsPath As String: PPMProjectsPath = PPMPaths.GetPPMProjectsPath()
    With NewFileSystemObject()
        Dim ProjectName As String
        ProjectName = .GetFileName(SelectedProject.Path)

        Dim ProjectTimeStamp As String
        ProjectTimeStamp = Strings.Format( _
            .GetFile(SelectedProject.Path).DateCreated, "ddmmyyyy_hhnnss" _
        )

        Dim FolderName As String
        FolderName = FString("{0}_{1}", ProjectName, ProjectTimeStamp)

        Dim ThisProjectPath As String
        ThisProjectPath = .BuildPath(PPMProjectsPath, FolderName)
        If Not .FolderExists(ThisProjectPath) Then FileSystem.MkDir ThisProjectPath
    End With

    GetProjectPath = ThisProjectPath
End Function


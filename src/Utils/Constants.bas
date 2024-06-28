Attribute VB_Name = "Constants"
'@Folder "PearPMProject.src.Utils"
Option Explicit

' CALCULATED CONSTANTS (PROPERTIES)
Public Property Get GlobalPPMPath() As String
    GlobalPPMPath = NewFileSystemObject().BuildPath(Interaction.Environ("APPDATA"), "ppm")
End Property

Public Property Get LocalPPMPath() As String
    LocalPPMPath = NewFileSystemObject().BuildPath(Interaction.Environ("LOCALAPPDATA"), "ppm")
End Property

Public Property Get LocalRegistryPath() As String
    LocalRegistryPath = NewFileSystemObject().BuildPath(LocalPPMPath, "registry")
End Property

Public Property Get ProjectConfigPath() As String
    On Error GoTo Catch
    If IsFalse(Constants.ProjectPath) Then Exit Property
    ProjectConfigPath = GetConfigFilePathFromFolder(Constants.ProjectPath)
    On Error GoTo 0
Exit Property

Catch:
    If Information.Err.Number = 76 Then
        Immediate.WriteLine "ERR: It is required to save the project before exporting."
        End
    Else
        Information.Err.Raise Information.Err.Number, "Constants"
    End If
End Property

Public Property Get UserConfigPath() As String
    UserConfigPath = GetConfigFilePathFromFolder(Interaction.Environ("USERPROFILE"))
End Property

Public Property Get GlobalConfigPath() As String
    If Not NewFileSystemObject().FolderExists(GlobalPPMPath) Then
        PFileSystem.CreateFolder GlobalPPMPath, Recoursive:=True
    End If
    GlobalConfigPath = GetConfigFilePathFromFolder(GlobalPPMPath)
End Property

Public Property Get ProjectPath() As String
    With NewFileSystemObject()
        Dim ProjectsPath As String
        ProjectsPath = .BuildPath(LocalPPMPath, "projects")
        Dim ProjectName As String
        ProjectName = .GetFileName(SelectedProject.Path)
        If IsFalse(ProjectName) Then Exit Property

        Dim ProjectTimeStamp As String
        ProjectTimeStamp = Strings.Format( _
            .GetFile(SelectedProject.Path).DateCreated, "ddmmyyyy_hhnnss" _
        )

        Dim FolderName As String
        FolderName = FString("{0}_{1}", ProjectName, ProjectTimeStamp)

        Dim ThisProjectPath As String
        ThisProjectPath = .BuildPath(ProjectsPath, FolderName)
        If Not .FolderExists(ThisProjectPath) Then PFileSystem.CreateFolder ThisProjectPath, Recoursive:=True
    End With

    ProjectPath = ThisProjectPath
End Property

Private Function GetConfigFilePathFromFolder(ByVal Folder As String) As String
    With NewFileSystemObject()
        GetConfigFilePathFromFolder = .BuildPath(Folder, ".ppmrc")
    End With
End Function

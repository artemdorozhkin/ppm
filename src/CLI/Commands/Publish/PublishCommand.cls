VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublishCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Publish"
Option Explicit

Implements ICommand

Private Type TExportCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "publish"
    this.CommandInfo.Description = ""
    this.CommandInfo.Usage = Array("[<options>]")
End Sub

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If Not SelectedProject.IsModuleExists("package") Then
        Immediate.WriteLine "for publish the package you need to run ppm ""init"""
        End
    End If

    Dim ms As Double: ms = DateTime.Timer
    Dim RootFolder As String: RootFolder = CreateExport()
    SelectedProject.ExportModule RootFolder, SelectedProject.GetModule("package")

    Dim PackInfo As Object: Set PackInfo = GetPackInfo()
    Dim ZipName As String
    ZipName = Me.GetBuildName(PackInfo) & ".zip"
    Dim ZipPath As String
    ZipPath = NewFileSystemObject().BuildPath(Interaction.Environ("TEMP"), ZipName)
    Dim Zip As PZip: Set Zip = NewPZip(ZipPath)

    Zip.AddFolderContent RootFolder
    FileSystem.Kill RootFolder & "\*"
    FileSystem.RmDir RootFolder

    Dim IsLocal As Boolean: IsLocal = this.Config.GetValue("to") = "local"
    If IsLocal Then PublishOnLocal ZipPath, PackInfo Else PublishOnServer ZipPath

    Immediate.WriteLine PStrings.FString( _
        "package '{0}' is published in {1}", PackInfo("name"), Utils.ConvertTime(DateTime.Timer - ms) _
    )
End Sub

Public Function GetBuildName(ByRef PackInfo As Object) As String
    GetBuildName = PStrings.FString("{0}_{1}", PackInfo("name"), Strings.Replace(PackInfo("version"), ".", "_"))
End Function

Public Function CreateExport() As String
    With NewFileSystemObject()
        Dim Folder As String
        Folder = .BuildPath(Interaction.Environ("TEMP"), .GetTempName())
        If Not .FolderExists(Folder) Then PFileSystem.CreateFoldersRecoursive Folder
    End With

    Dim Exporter As ProjectExporter
    Set Exporter = NewProjectExporter( _
        Destination:=Folder, _
        RewriteLastExport:=False _
    )
    Exporter.ExportFolders "UTF-8", SelectedProject.Name

    CreateExport = Folder
End Function

Public Function GetPackInfo() As Object
    Dim Module As VBComponent
    Set Module = SelectedProject.GetModule("package")
    Dim Code As String
    Code = Utils.UncommentString(Module.CodeModule.Lines(1, Module.CodeModule.CountOfLines))

    Set GetPackInfo = PJSON.Parse(Code)
End Function

Public Function PublishOnLocal(ByVal ZipPath As String, ByRef PackInfo As Object)
    Dim LocalRegistryPath As String
    LocalRegistryPath = this.Config.GetValue("registry")

    With NewFileSystemObject()
        Dim PackageFolder As String
        PackageFolder = .BuildPath(LocalRegistryPath, SelectedProject.Name)
        If Not .FolderExists(PackageFolder) Then PFileSystem.CreateFoldersRecoursive PackageFolder
        Dim BuildZip As String
        BuildZip = .BuildPath(PackageFolder, Me.GetBuildName(PackInfo) & ".zip")
    End With

    FileSystem.FileCopy ZipPath, BuildZip
    FileSystem.Kill ZipPath

    Me.CreateManifest PackageFolder, PackInfo
End Function

Public Function PublishOnServer(ByVal ZipPath As String)
    
End Function

Public Sub CreateManifest(ByVal PackageFolder As String, ByRef PackInfo As Object)
    With NewFileSystemObject()
        Dim ManifestPath As String
        ManifestPath = .BuildPath(PackageFolder, "manifest.json")
        If .FileExists(ManifestPath) Then Exit Sub
    End With

    Dim JB As PJSONBuilder: Set JB = New PJSONBuilder
    JB.AddObject _
        "name", PackInfo("name"), _
        "latest", PackInfo("version"), _
        "buildPath", PStrings.FString("{0}\{1}.zip", PackageFolder, Me.GetBuildName(PackInfo))

    PFileSystem.SaveToFile ManifestPath, JB.ToString()
End Sub


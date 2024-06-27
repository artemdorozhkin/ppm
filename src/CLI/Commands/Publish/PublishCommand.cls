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
    CLI.Lang.SetBlock "/root/ppmCommands/publish"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "publish"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("local") = CLI.Lang.GetValue("param", "name=local")
    this.CommandInfo.Params.Item("registry") = CLI.Lang.GetValue("param", "name=registry")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
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
    If Not SelectedProject.IsComponentExists("package") Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/packageModuleNotFound"))
        End
    End If

    Immediate.WriteLine CLI.Lang.GetValue("messages/publishing")
    Dim ms As Double: ms = DateTime.Timer

    MoveProjectToPPM
    Dim RootFolder As String: RootFolder = CreateExport()
    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent("package"))
    Pack.RenameModule Pack.Name & "Package"
    SelectedProject.ExportComponent RootFolder, Pack.Component
    Pack.RenameModule "package"

    Dim ZipName As String
    ZipName = Me.GetBuildName(Pack) & ".zip"
    Dim ZipPath As String
    ZipPath = NewFileSystemObject().BuildPath(Interaction.Environ("TEMP"), ZipName)
    Dim Zip As PZip: Set Zip = NewPZip(ZipPath)
    Zip.AddFolderContent RootFolder
    FileSystem.Kill RootFolder & "\*"
    FileSystem.RmDir RootFolder

    Dim ToLocal As Boolean: ToLocal = this.Config.GetValue("local")
    Dim Registry As String: Registry = this.Config.GetValue("registry")
    If ToLocal Then
        If Not NewFileSystemObject().FolderExists(Registry) Then
            Registry = Constants.LocalRegistryPath
        End If
    End If

    Dim Success As Boolean
    If ToLocal Then
        PublishOnLocal Registry, ZipPath, Pack
        Success = True
    Else
        Dim Response As TResponse
        Response = PublishOnServer(Registry, ZipPath, Pack)
        Success = Response.Code = HTTPCodes.OK_200 Or _
                  Response.Code = HTTPCodes.CREATED_201
    End If

    MovePPMToProject

    If Success Then
        Immediate.WriteLine PStrings.FString( _
            CLI.Lang.GetValue("messages/published"), Pack.Name, Utils.ConvertTime(DateTime.Timer - ms) _
        )
    Else
        Immediate.WriteLine CLI.Lang.GetValue("messages/somethingWentWrong")
        Immediate.WriteLine Response.Text
    End If
End Sub

Public Sub MoveProjectToPPM()
    Dim Component As VBComponent
    For Each Component In SelectedProject.Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, SelectedProject.Name) Then
            Module.Move "ppm-modules." & Strings.Replace(Module.Path, "Project", "", Count:=1)
        End If
    Next
End Sub

Public Sub MovePPMToProject()
    Dim Component As VBComponent
    For Each Component In SelectedProject.Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, "ppm-modules") Then
            Dim PackageName As String
            PackageName = Strings.Replace(SelectedProject.Name, "Project", "")
            Dim CorrectPath As String
            CorrectPath = Strings.Replace(Module.Path, PackageName, SelectedProject.Name, Count:=1)
            CorrectPath = Strings.Replace(CorrectPath, "ppm-modules.", "")
            Module.Move CorrectPath
        End If
    Next
End Sub

Public Function GetBuildName(ByRef Pack As Pack) As String
    GetBuildName = PStrings.FString("{0}-{1}", Pack.Name, Pack.Version)
End Function

Public Function CreateExport() As String
    With NewFileSystemObject()
        Dim Folder As String
        Folder = .BuildPath(Interaction.Environ("TEMP"), .GetTempName())
        If Not .FolderExists(Folder) Then PFileSystem.CreateFolder Folder, Recoursive:=True
    End With

    Dim Exporter As ProjectExporter
    Set Exporter = NewProjectExporter( _
        Destination:=Folder, _
        RewriteLastExport:=False _
    )
    Exporter.ExportFolders "UTF-8", "ppm-modules"

    CreateExport = Folder
End Function

Public Function PublishOnLocal(ByVal Registry As String, ByVal ZipPath As String, ByRef Pack As Pack) As Boolean
    With NewFileSystemObject()
        Dim PackageFolder As String
        PackageFolder = .BuildPath(Registry, Pack.Name)
        If Not .FolderExists(PackageFolder) Then PFileSystem.CreateFolder PackageFolder, Recoursive:=True
        Dim BuildZip As String
        BuildZip = .BuildPath(PackageFolder, Me.GetBuildName(Pack) & ".zip")
    End With

    FileSystem.FileCopy ZipPath, BuildZip
    FileSystem.Kill ZipPath

    Pack.JSON("packUrl") = BuildZip
    Me.CreateOrUpdateManifest PackageFolder, Pack
    PublishOnLocal = True
End Function

Public Function PublishOnServer(ByVal Registry As String, ByVal ZipPath As String, ByRef Pack As Pack) As TResponse
    Dim FormData As FormData: Set FormData = New FormData
    Dim Boundary As String: Boundary = FormData.GenerateBoundary()

    Dim PackJSON As Object: Set PackJSON = Pack.JSON
    PackJSON("checkSum") = Utils.CalculateFileCheckSum(ZipPath)
    Dim PackStr As String
    PackStr = PJSON.Stringify(PackJSON)
    Dim Body() As Byte
    Body = FormData.Create(Boundary) _
                   .AddFile(ZipPath, "package", "application/zip") _
                   .AddText(PackStr, "packInfo") _
                   .ToBytes()

    Dim Headers As Object: Set Headers = NewDictionary()
    Headers("Content-Type") = "multipart/form-data; boundary=" & Boundary

    Dim HTTP As HTTP: Set HTTP = New HTTP
    PublishOnServer = HTTP.PostRequest( _
        Registry, Headers, Body _
    )

    FileSystem.Kill ZipPath
End Function

Public Sub CreateOrUpdateManifest(ByVal PackageFolder As String, ByRef Pack As Pack)
    Dim JB As PJSONBuilder: Set JB = New PJSONBuilder
    With NewFileSystemObject()
        Dim ManifestPath As String
        ManifestPath = .BuildPath(PackageFolder, "manifest.json")
        If .FileExists(ManifestPath) Then
            JB.SetJSON PJSON.Parse(PFileSystem.ReadFile(ManifestPath))
        Else
            JB.AddObject _
                "name", Pack.Name, _
                "latest", "", _
                "versions", NewDictionary()
        End If
    End With
    Dim Manifest As Manifest: Set Manifest = NewManifest(JB.ToString())
    Set Manifest.Versions(Pack.Version) = Pack.JSON
    Manifest.Latest = GetLatest(Manifest.Versions)

    PFileSystem.SaveToFile ManifestPath, Manifest.ToString()
End Sub

Public Function GetLatest(ByRef Versions As Object) As String
    Dim Latest As String
    Dim Vers As Variant
    For Each Vers In Versions
        Latest = Utils.GetNewest(Vers, Latest)
    Next

    GetLatest = Latest
End Function

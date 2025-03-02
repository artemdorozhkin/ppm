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
    Dim Config As Config
    Set Config = NewConfig(ConfigScopes.UserScope)

    Dim Credentials As Object
    Set Credentials = NewDictionary()

    Credentials("email") = Config.GetValue("email")
    Credentials("api_key") = Config.GetValue("api_key")

    If Not SelectedProject.IsComponentExists("package") Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/packageModuleNotFound"))
        End
    End If

    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent("package"))
    Dim PackRoot As String: PackRoot = Pack.Root
    Immediate.WriteLine PStrings.FString("{0}@{1}", Pack.Name, Pack.Version)
    Immediate.WriteLine CLI.Lang.GetValue("messages/publishing")
    Dim ms As Double: ms = DateTime.Timer

    PrintFilesList PackRoot

    MoveProjectToPPM PackRoot, Pack.Name
    Pack.RenameComponent Pack.Name & "Package"
    Dim RootFolder As String: RootFolder = CreateExport()
    SelectedProject.ExportComponent RootFolder, Pack.Component
    Pack.RenameComponent "package"

    Dim ZipName As String
    ZipName = Me.GetBuildName(Pack) & ".zip"
    Dim ZipPath As String
    ZipPath = NewFileSystemObject().BuildPath(Interaction.Environ("TEMP"), ZipName)
    Dim Zip As PZip: Set Zip = NewPZip(ZipPath)
    Zip.AddFolderContent RootFolder
    FileSystem.Kill RootFolder & "\*"
    FileSystem.RmDir RootFolder

    PrintDetails Zip, Pack

    Dim ToLocal As Boolean
    ToLocal = GetFirstValueFrom("local", this.Tokens, this.Config)
    Dim Registry As String
    Registry = GetFirstValueFrom("registry", this.Tokens, this.Config)
    Registry = Utils.ResolveUrl(Registry)
    If ToLocal Then
        If Not NewFileSystemObject().FolderExists(Registry) Then
            PFileSystem.CreateFolder Registry, Recoursive:=True
        End If
    End If

    Dim Success As Boolean
    If ToLocal Then
        Success = PublishOnLocal(Registry, ZipPath, Pack)
    Else
        Dim Response As TResponse
        Response = PublishOnServer(Registry, ZipPath, Pack, Credentials)
        Success = Response.Code = HTTPCodes.OK_200 Or _
                  Response.Code = HTTPCodes.CREATED_201
    End If

    MovePPMToProject PackRoot, Pack.Name

    Immediate.WriteLine
    If Success Then
        Immediate.WriteLine PStrings.FString( _
            CLI.Lang.GetValue("messages/published"), Pack.Name, Utils.ConvertTime(DateTime.Timer - ms) _
        )
    Else
        Immediate.WriteLine CLI.Lang.GetValue("messages/somethingWentWrong")
        Immediate.WriteLine Response.Text
    End If
End Sub

Public Sub MoveProjectToPPM(ByVal ProjectRoot As String, ByVal PackName As String)
  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    For Each Component In SelectedProject.Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, ProjectRoot) Then
            Module.Move "ppm-modules." & Strings.Replace(Module.Path, ProjectRoot, PackName, Count:=1)
        End If
    Next
End Sub

Public Sub MovePPMToProject(ByVal ProjectRoot As String, ByVal PackName As String)
  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    For Each Component In SelectedProject.Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, "ppm-modules") Then
            Dim CorrectPath As String
            CorrectPath = Strings.Replace(Module.Path, PackName, ProjectRoot, Count:=1)
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

    Dim Sep As String: Sep = Application.PathSeparator
    Pack.JSON("packUrl") = Strings.Replace(BuildZip, Sep, Sep & Sep)
    Me.CreateOrUpdateManifest PackageFolder, Pack
    PublishOnLocal = True
End Function

Public Function PublishOnServer(ByVal Registry As String, ByVal ZipPath As String, ByRef Pack As Pack, ByRef Credentials As Object) As TResponse
    Dim FormData As FormData: Set FormData = New FormData

  #If DEV Then
    Dim PackJSON As Dictionary: Set PackJSON = Pack.JSON
  #Else
    Dim PackJSON As Object: Set PackJSON = Pack.JSON
  #End If
    PackJSON("checkSum") = Utils.CalculateFileCheckSum(ZipPath)
    Dim PackStr As String
    PackStr = PJSON.Stringify(PackJSON)
    Dim Body() As Byte
    Body = FormData.Create() _
                   .AddFile(ZipPath, "package", "application/zip") _
                   .AddText(PackStr, "packInfo") _
                   .ToBytes()
  #If DEV Then
    Dim Headers As Dictionary: Set Headers = NewDictionary()
  #Else
    Dim Headers As Object: Set Headers = NewDictionary()
  #End If
    Headers("Content-Type") = "multipart/form-data; boundary=" & FormData.Boundary
    Headers("email") = Credentials("email")
    Headers("api-key") = Credentials("api_key")

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

#If DEV Then
  Public Function GetLatest(ByRef Versions As Dictionary) As String
#Else
  Public Function GetLatest(ByRef Versions As Object) As String
#End If
    Dim Latest As String
    Dim Vers As Variant
    For Each Vers In Versions
        Latest = Utils.GetNewestVersion(Vers, Latest)
    Next

    GetLatest = Latest
End Function

Public Sub PrintFilesList(ByVal ProjectRoot As String)
    Immediate.WriteLine
    Immediate.WriteLine CLI.Lang.GetValue("messages/filesToPublish")

  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    For Each Component In SelectedProject.Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, ProjectRoot) Then
            Dim Path As String
            Path = Strings.Replace(Module.Path & "." & Module.Name, ProjectRoot, "")
            Path = Strings.Replace(Path, ".", Application.PathSeparator)
            Immediate.WriteLine Path & Module.Ext
        End If
    Next
End Sub

Public Sub PrintDetails(ByRef Zip As PZip, ByRef Pack As Pack)
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/header")
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/name") & Pack.Name
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/version") & Pack.Version
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/packSize") & Utils.SizeInStr(Zip.Size)
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/unpackedSize") & Utils.SizeInStr(Zip.UnpackedSize)
    Immediate.WriteLine CLI.Lang.GetValue("messages/packDetails/totalFiles") & Zip.FilesCount
End Sub

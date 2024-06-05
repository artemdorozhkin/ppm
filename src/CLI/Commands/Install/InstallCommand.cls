VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstallCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Install"
Option Explicit

Implements ICommand

Private Type TInstallCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
    Registry As String
    FromLocal As Boolean
    Parents As Object
End Type

Private this As TInstallCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "install"
    this.CommandInfo.Description = "Installs the selected packages."
    this.CommandInfo.Params.Item("local") = "\\tInstall packages and dependencies from local registry."
    this.CommandInfo.Params.Item("registry") = "\\tSpecific registry path or url."
    this.CommandInfo.Usage = Array("[<some package>@<some version>] [<options>]")
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
    Set this.Parents = NewDictionary(vbTextCompare)
    this.Registry = this.Config.GetValue("registry")
    this.FromLocal = this.Config.GetValue("local")
    If this.FromLocal Then
        If Not NewFileSystemObject().FolderExists(this.Registry) Then
            this.Registry = Constants.LocalRegistryPath
        End If
    End If
    Dim ms As Double: ms = DateTime.Timer
    Immediate.WriteLine "installing..."
    Dim Libs As Object: Set Libs = Me.GetFlattedLibs()

    Dim Paths As Variant
    If this.FromLocal Then
        Paths = Me.GetLocalLibs(Libs)
    Else
        Paths = Me.DownloadLibs(Libs)
    End If
    InstallLibs Paths
    UpdatePack

    If Libs.Count = 0 Then
        Immediate.WriteLine "libs to install not found"
    Else
        ms = DateTime.Timer - ms
        Immediate.WriteLine PStrings.FString( _
            "successfully installed {0} libs in {1}", _
            Libs.Count, _
            Utils.ConvertTime(ms) _
        )
    End If
End Sub

Public Function GetFlattedLibs() As Object
    If this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        Set GetFlattedLibs = Me.GetLibsFromTokens()
    Else
        Set GetFlattedLibs = Me.GetLibsFromDependencies()
    End If
End Function

Public Function GetLibsFromTokens() As Object
    Dim Libs As Dictionary: Set Libs = NewDictionary(vbTextCompare)
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind <> TokenKind.Identifier Then GoTo Continue
        Dim LibName As String: LibName = Token.Text
        Dim Version As String: Version = "latest"
        If PStrings.Contains(LibName, "@") Then
            Dim Parts As Variant: Parts = Strings.Split(LibName, "@")
            Version = Parts(1)
            LibName = Parts(0)
        End If
        Dim Manifest As Object
        If this.FromLocal Then
            Set Manifest = Me.GetLocalManifest(LibName, Version)
        Else
            Set Manifest = Me.GetManifestFromServer(LibName, Version)
        End If
        this.Parents.Add Manifest("name"), Manifest("version")
        AddFromManifest Manifest, Libs
Continue:
    Next

    Set GetLibsFromTokens = Libs
End Function

Public Function GetLibsFromDependencies() As Object
    Dim Libs As Dictionary: Set Libs = NewDictionary(vbTextCompare)
    Dim Dependencies As Object: Set Dependencies = NewPack(SelectedProject).Read().Dependencies
    If IsFalse(Dependencies) Then
        Set GetLibsFromDependencies = Libs
        Exit Function
    End If

    If Dependencies.Count = 0 Then
        Set GetLibsFromDependencies = Libs
        Exit Function
    End If

    Dim LibName As Variant
    For Each LibName In Dependencies
        Dim Version As String: Version = Dependencies(LibName)
        this.Parents.Add LibName, Version
        If this.FromLocal Then
            AddFromManifest Me.GetLocalManifest(LibName, Dependencies(LibName)), Libs
        Else
            AddFromManifest Me.GetManifestFromServer(LibName, Version), Libs
        End If
Continue:
    Next

    Set GetLibsFromDependencies = Libs
End Function

Public Sub AddFromManifest(ByRef Manifest As Object, ByRef Libs As Object)
    Dim LibName As String: LibName = Manifest("name")
    Dim Version As String: Version = Manifest("version")
    Dim Dependencies As Object
    If Manifest.Exists("dependencies") Then
        Set Dependencies = Manifest("dependencies")
    End If

    If Libs.Exists(LibName) Then
        Dim Newest As String
        Newest = Utils.GetNewest(Libs(LibName), Version)
        Libs(LibName) = Newest
    Else
        Libs.Add LibName, Version
    End If

    If this.Parents.Exists(LibName) Then
        this.Parents(LibName) = Utils.GetNewest(this.Parents(LibName), Version)
    End If

    If IsFalse(Dependencies) Then Exit Sub

    Dim DepName As Variant
    For Each DepName In Dependencies.Keys()
        Newest = Utils.GetNewest(Dependencies(DepName), Libs(DepName))
        If Not Libs.Exists(DepName) Or Libs(DepName) <> Newest Then
            If this.FromLocal Then
                AddFromManifest Me.GetLocalManifest(DepName, Dependencies(DepName)), Libs
            Else
                AddFromManifest Me.GetManifestFromServer(DepName, Dependencies(DepName)), Libs
            End If
        End If
    Next
End Sub

Public Function GetLocalManifest(ByVal LibName As String, ByVal Version As String) As Object
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim ManifestPath As String
    ManifestPath = FSO.BuildPath(this.Registry, LibName)
    ManifestPath = FSO.BuildPath(ManifestPath, "manifest.json")

    Dim JSONString As String: JSONString = PFileSystem.ReadFile(ManifestPath)
    Dim Manifest As Object: Set Manifest = PJSON.Parse(JSONString)
    If PStrings.IsEqual(Version, "latest") Then
        Version = Manifest("latest")
    End If

    Set GetLocalManifest = Manifest("versions")(Version)
End Function

Public Function GetManifestFromServer(ByVal LibName As String, ByVal Version As String) As Object
    Dim URL As String
    If Strings.Right(this.Registry, 1) <> "/" Then
        URL = this.Registry & PStrings.FString("/{0}/{1}", LibName, Version)
    Else
        URL = this.Registry & PStrings.FString("{0}/{1}", LibName, Version)
    End If

    Dim HTTP As HTTP: Set HTTP = New HTTP
    Dim Response As TResponse
    Response = HTTP.GetRequest(URL)
    If Response.Code <> HTTPCodes.OK_200 Then
        Immediate.WriteLine "ERR: something went wrong, try again later"
        Immediate.WriteLine Response.Text
        End
    End If

    Set GetManifestFromServer = PJSON.Parse(Response.Text)
End Function

Public Function GetLocalLibs(ByRef Libs As Dictionary) As Variant
    Dim Paths As Collection: Set Paths = New Collection
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim CacheFolder As String
    CacheFolder = FSO.BuildPath(Constants.LocalPPMPath, "cache")
    PFileSystem.CreateFolder CacheFolder, Recoursive:=True

    Dim Name As Variant
    For Each Name In Libs
        Dim Version As String: Version = Libs(Name)

        Dim FileName As String
        FileName = PStrings.FString("{0}-{1}.zip", Name, Version)
        Dim Path As String
        Path = FSO.BuildPath(this.Registry, Name)
        Path = FSO.BuildPath(Path, FileName)

        Dim SavePath As String
        SavePath = FSO.BuildPath(CacheFolder, FileName)
        On Error Resume Next
        FileSystem.FileCopy Path, SavePath
        Paths.Add SavePath
    Next
    On Error GoTo 0

    GetLocalLibs = PCollection.ToArray(Paths)
End Function

Public Function DownloadLibs(ByRef Libs As Dictionary) As Variant
    Dim Registry As String
    If Strings.Right(this.Registry, 1) <> "/" Then
        Registry = this.Registry & "/"
    End If

    Dim HTTP As HTTP: Set HTTP = New HTTP
    Dim Converter As BinaryConverter: Set Converter = New BinaryConverter
    Dim Paths As Collection: Set Paths = New Collection
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim CacheFolder As String
    CacheFolder = FSO.BuildPath(Constants.LocalPPMPath, "cache")
    PFileSystem.CreateFolder CacheFolder, Recoursive:=True

    Dim Name As Variant
    For Each Name In Libs
        Dim Version As String: Version = Libs(Name)
        Dim URL As String
        URL = PStrings.FString("{0}{1}/{2}", Registry, Name, Version)
        Dim Response As TResponse
        Response = HTTP.GetRequest(URL)
        If Response.Code <> HTTPCodes.OK_200 Then
            Information.Err.Raise _
                17, _
                Information.TypeName(Me), _
                "ERR: download failed: " & PStrings.FString("{0}@{1}\\n{2}", Name, Version, Response.Text)
        End If

        Dim Manifest As Object: Set Manifest = PJSON.Parse(Response.Text)
        URL = Manifest("packUrl")
        Response = HTTP.GetRequest(URL)
        If Response.Code <> HTTPCodes.OK_200 Then
            Information.Err.Raise _
                17, _
                Information.TypeName(Me), _
                "ERR: download failed: " & PStrings.FString("{0}@{1}\\n{2}", Name, Version, Response.Text)
        End If

        Dim Parts As Variant: Parts = Strings.Split(URL, "/")
        Dim FileName As String
        FileName = Parts(UBound(Parts))
        Dim SavePath As String
        SavePath = FSO.BuildPath(CacheFolder, FileName)

        If Not FSO.FileExists(SavePath) Then
            Converter.ToFile Response.Body, SavePath
        End If
        Paths.Add SavePath
    Next

    DownloadLibs = PCollection.ToArray(Paths)
End Function

Public Sub InstallLibs(ByRef Paths As Variant)
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim Zip As PZip: Set Zip = New PZip
    Dim Path As Variant
    For Each Path In Paths
        If Not FSO.FileExists(Path) Then GoTo Continue
        Zip.SetZip Path
        Dim TempFolder As String
        TempFolder = FSO.BuildPath(FSO.GetParentFolderName(Path), FSO.GetTempName())
        Zip.Unzip TempFolder
        FileSystem.Kill Path
        Me.ImportModules TempFolder
        FileSystem.RmDir TempFolder
Continue:
    Next
End Sub

Public Sub ImportModules(ByVal Path As String)
    Dim Folder As Folder: Set Folder = NewFolder(Path)
    Dim File As File
    For Each File In Folder.Files
        On Error Resume Next
        SelectedProject.ImportModule File.Path
        File.Delete Force:=True
    Next
End Sub

Public Sub UpdatePack()
    NewPack(SelectedProject).Read().Update Dependencies:=this.Parents
End Sub

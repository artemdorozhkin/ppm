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
  #If DEV Then
    Parents As Dictionary
  #Else
    Parents As Object
  #End If
End Type

Private this As TInstallCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/install"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "install"
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
    Set this.Parents = NewDictionary(vbTextCompare)
    this.Registry = this.Config.GetValue("registry")
    this.FromLocal = this.Config.GetValue("local")
    If this.FromLocal Then
        If Not NewFileSystemObject().FolderExists(this.Registry) Then
            this.Registry = Constants.LocalRegistryPath
        End If
    End If
    Dim ms As Double: ms = DateTime.Timer
    Immediate.WriteLine CLI.Lang.GetValue("messages/installing")
  #If DEV Then
    Dim Libs As Dictionary: Set Libs = Me.GetFlattedLibs()
  #Else
    Dim Libs As Object: Set Libs = Me.GetFlattedLibs()
  #End If

    Dim Paths As Variant
    If this.FromLocal Then
        Paths = Me.GetLocalLibs(Libs)
    Else
        Paths = Me.DownloadLibs(Libs)
    End If
    InstallLibs Paths
    UpdatePack

    If Libs.Count = 0 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/libsNotFound")
    Else
        ms = DateTime.Timer - ms
        Immediate.WriteLine PStrings.FString( _
            CLI.Lang.GetValue("messages/installed"), _
            Libs.Count, _
            Utils.ConvertTime(ms) _
        )
    End If
End Sub

#If DEV Then
  Public Function GetFlattedLibs() As Dictionary
#Else
  Public Function GetFlattedLibs() As Object
#End If
    If this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        Set GetFlattedLibs = Me.GetLibsFromTokens()
    Else
        Set GetFlattedLibs = Me.GetLibsFromDependencies()
    End If
End Function

#If DEV Then
  Public Function GetLibsFromTokens() As Dictionary
#Else
  Public Function GetLibsFromTokens() As Object
#End If
  #If DEV Then
    Dim Libs As Dictionary: Set Libs = NewDictionary(vbTextCompare)
  #Else
    Dim Libs As Object: Set Libs = NewDictionary(vbTextCompare)
  #End If
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
        Dim LibPack As Pack
        If this.FromLocal Then
            Set LibPack = Me.GetLocalLibPack(LibName, Version)
        Else
            Set LibPack = Me.GetLibPackFromServer(LibName, Version)
        End If
        this.Parents.Add LibPack.Name, LibPack.Version
        AddFromLibPack LibPack, Libs
Continue:
    Next

    Set GetLibsFromTokens = Libs
End Function

#If DEV Then
  Public Function GetLibsFromDependencies() As Dictionary
#Else
  Public Function GetLibsFromDependencies() As Object
#End If
  #If DEV Then
    Dim Libs As Dictionary: Set Libs = NewDictionary(vbTextCompare)
    Dim Dependencies As Dictionary
    Set Dependencies = NewPack(SelectedProject.GetComponent("package")).Dependencies
  #Else
    Dim Libs As Object: Set Libs = NewDictionary(vbTextCompare)
    Dim Dependencies As Object
    Set Dependencies = NewPack(SelectedProject.GetComponent("package")).Dependencies
  #End If
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
            AddFromLibPack Me.GetLocalLibPack(LibName, Dependencies(LibName)), Libs
        Else
            AddFromLibPack Me.GetLibPackFromServer(LibName, Version), Libs
        End If
    Next

    Set GetLibsFromDependencies = Libs
End Function

#If DEV Then
  Public Sub AddFromLibPack(ByRef LibPack As Pack, ByRef Libs As Dictionary)
#Else
  Public Sub AddFromLibPack(ByRef LibPack As Pack, ByRef Libs As Object)
#End If
    Dim LibName As String: LibName = LibPack.Name
    Dim Version As String: Version = LibPack.Version
  #If DEV Then
      Dim Dependencies As Dictionary
  #Else
      Dim Dependencies As Object
  #End If
    If IsTrue(LibPack.Dependencies) Then
        Set Dependencies = LibPack.Dependencies
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
                AddFromLibPack Me.GetLocalLibPack(DepName, Dependencies(DepName)), Libs
            Else
                AddFromLibPack Me.GetLibPackFromServer(DepName, Dependencies(DepName)), Libs
            End If
        End If
    Next
End Sub

Public Function GetLocalLibPack(ByVal LibName As String, ByVal Version As String) As Pack
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim ManifestPath As String
    ManifestPath = FSO.BuildPath(this.Registry, LibName)
    ManifestPath = FSO.BuildPath(ManifestPath, "manifest.json")
    If Not FSO.FileExists(ManifestPath) Then
        Information.Err.Raise _
            76, _
            Information.TypeName(Me), _
            PStrings.FString(CLI.Lang.GetValue("messages/manifestNotFound"), LibName)
    End If

    Dim JSONString As String: JSONString = PFileSystem.ReadFile(ManifestPath)
    Dim Manifest As Manifest: Set Manifest = NewManifest(JSONString)
    If PStrings.IsEqual(Version, "latest") Then
        Version = Manifest.Latest
    End If

    Set GetLocalLibPack = NewPack(Manifest.Versions(Version))
End Function

Public Function GetLibPackFromServer(ByVal LibName As String, ByVal Version As String) As Pack
    Dim URL As String
    If Strings.Right(this.Registry, 1) <> "/" Then
        URL = this.Registry & PStrings.FString("/{0}/{1}", LibName, Version)
    Else
        URL = this.Registry & PStrings.FString("{0}/{1}", LibName, Version)
    End If

    Dim HTTP As HTTP: Set HTTP = New HTTP
    Dim Response As TResponse
    On Error GoTo Catch
    Response = HTTP.GetRequest(URL)
    On Error GoTo 0
    If Response.Code <> HTTPCodes.OK_200 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/somethingWentWrong")
        Immediate.WriteLine Response.Text
        End
    End If

    Set GetLibPackFromServer = NewPack(PJSON.Parse(Response.Text))
Exit Function

Catch:
    Const WRONG_URL_ERROR As Long = &H800C0005
    If Information.Err.Number = WRONG_URL_ERROR Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/wrongURL"), URL)
        End
    Else
        With Information.Err
            .Raise .Number, .Source, .Description
        End With
    End If
End Function

#If DEV Then
  Public Function GetLocalLibs(ByRef Libs As Dictionary) As Variant
#Else
  Public Function GetLocalLibs(ByRef Libs As Object) As Variant
#End If
    Dim Paths As Collection: Set Paths = New Collection
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim CacheFolder As String
    CacheFolder = FSO.BuildPath(Constants.LocalPPMPath, "cache")
    PFileSystem.CreateFolder CacheFolder, Recoursive:=True

    ' TODO: use manifest packUrl (don't forget replace escaped \\ in path)
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

#If DEV Then
  Public Function DownloadLibs(ByRef Libs As Dictionary) As Variant
#Else
  Public Function DownloadLibs(ByRef Libs As Object) As Variant
#End If
    Dim Registry As String
    If Strings.Right(this.Registry, 1) <> "/" Then
        Registry = this.Registry & "/"
    End If

    Dim HTTP As HTTP: Set HTTP = New HTTP
    Dim Converter As BinaryConverter: Set Converter = New BinaryConverter
    Dim Paths As Collection: Set Paths = New Collection
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
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
                PStrings.FString( _
                    CLI.Lang.GetValue("messages/downloadFailed"), _
                    PStrings.FString("{0}@{1}\\n{2}", Name, Version, Response.Text) _
                )
        End If

      #If DEV Then
        Dim Manifest As Dictionary: Set Manifest = PJSON.Parse(Response.Text)
      #Else
        Dim Manifest As Object: Set Manifest = PJSON.Parse(Response.Text)
      #End If
        URL = Manifest("packUrl")
        Response = HTTP.GetRequest(URL)
        If Response.Code <> HTTPCodes.OK_200 Then
            Information.Err.Raise _
                17, _
                Information.TypeName(Me), _
                PStrings.FString( _
                    CLI.Lang.GetValue("messages/downloadFailed"), _
                    PStrings.FString("{0}@{1}\\n{2}", Name, Version, Response.Text) _
                )
        End If

        If Not PStrings.IsEqual( _
            Manifest("checkSum"), _
            Utils.CalculateBytesCheckSum(Response.Body) _
        ) Then
            Information.Err.Raise _
                59, _
                Information.TypeName(Me), _
                CLI.Lang.GetValue("messages/invalidCheckSum")
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
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim Zip As PZip: Set Zip = New PZip
    Dim Path As Variant
    For Each Path In Paths
        If Not FSO.FileExists(Path) Then GoTo Continue
        Zip.SetZip Path
        Dim TempFolder As String
        TempFolder = FSO.BuildPath(FSO.GetParentFolderName(Path), FSO.GetTempName())
        Zip.Unzip TempFolder
        FileSystem.Kill Path
        Me.ImportComponents TempFolder
        FileSystem.RmDir TempFolder
Continue:
    Next
End Sub

Public Sub ImportComponents(ByVal Path As String)
  #If DEV Then
    Dim Folder As Scripting.Folder: Set Folder = NewFolder(Path)
    Dim File As File
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim Folder As Object: Set Folder = NewFolder(Path)
    Dim File As Object
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    For Each File In Folder.Files
        On Error Resume Next
        Dim Ext As String: Ext = FSO.GetExtensionName(File.Name)
        Dim Name As String: Name = FSO.GetBaseName(File.Name)
        Dim CodePage As String: CodePage = Locale.GetAnsiCodePage()
        Dim LocalePath As String
        LocalePath = FSO.BuildPath( _
            FSO.GetParentFolderName(File.Path), _
            PStrings.FString("{0}{1}.{2}", Name, CodePage, Ext) _
        )
        PFileSystem.SaveToFile _
            LocalePath, _
            PFileSystem.ReadFile(File.Path), _
            PStrings.FString("Windows-{0}", CodePage)

        SelectedProject.ImportComponent LocalePath
        File.Delete Force:=True
        FileSystem.Kill LocalePath
    Next
End Sub

Public Sub UpdatePack()
    NewPack(SelectedProject.GetComponent("package")).Update Dependencies:=this.Parents
End Sub

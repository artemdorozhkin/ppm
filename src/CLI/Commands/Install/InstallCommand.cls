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
End Type

Private this As TInstallCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "install"
    this.CommandInfo.Description = "Installs the selected package."
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
    Dim Libs As Object: Set Libs = Me.GetFlattedLibs()
    Dim Registry As String: Registry = this.Config.GetValue("registry")
    Dim Paths As Variant
    If NewFileSystemObject().FolderExists(Registry) Then
        Paths = Me.GetLocalLibs(Registry, Libs)
    Else
        Paths = Me.DownloadLibs(Registry, Libs)
    End If
    InstallLibs Paths
End Sub

Public Function GetFlattedLibs() As Object
    If Not this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        Set GetFlattedLibs = Me.GetDependencies()
        Exit Function
    End If

    Dim Libs As Dictionary: Set Libs = NewDictionary(vbTextCompare)
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind <> TokenKind.Identifier Then GoTo Continue
        Me.FlatLib Libs, Token.Text
Continue:
    Next

    Set GetFlattedLibs = Libs
End Function

Public Function GetDependencies() As Object
    Dim Pack As Pack: Set Pack = NewPack(SelectedProject)
    If this.Tokens.IncludeDefinition(Definitions("dev-dep")) Then
        Set GetDependencies = Me.FlatDependencies(Pack.DevDependencies)
    Else
        Set GetDependencies = Me.FlatDependencies(Pack.Dependencies)
    End If
End Function

Public Function FlatDependencies(ByRef Dependencies As Dictionary) As Dictionary
    Dim Flatted As Dictionary: Set Flatted = NewDictionary(vbTextCompare)
    Dim Name As Variatn
    For Each Name In Dependencies
        Me.FlatLib Dependencies, PStrings("{0}@{1}", Name, Dependencies(Name))
    Next

    Set FlatDependencies = Dependencies
End Function

Public Sub FlatLib(ByRef Libs As Dictionary, ByVal Lib As String)
    Dim LibName As String
    Dim LibVersion As String
    If PStrings.Contains(Lib, "@") Then
        Dim NameVersion As Variant: NameVersion = Strings.Split(Lib, "@")
        LibName = NameVersion(0)
        LibVersion = NameVersion(1)
    Else
        LibName = Lib
        LibVersion = "latest"
    End If

    If Libs.Exists(LibName) Then
        LibVersion = GetNewest(Libs(LibName), LibVersion)
        Libs(LibName) = LibVersion
    Else
        Libs.Add LibName, LibVersion
    End If
End Sub

Public Function GetNewest(ByVal Current As String, ByVal Other As String) As Boolean
    Const Major As Integer = 0
    Const Minor As Integer = 1
    Const Patch As Integer = 2

    If PStrings.IsEqual(Current, "latest") Then
        GetNewest = Current
        Exit Function
    End If

    Dim CurrentParts As Variant: CurrentParts = Strings.Split(Current, ".")
    Dim OtherParts As Variant: OtherParts = Strings.Split(Other, ".")
    CurrentParts(Major) = Conversion.CInt(CurrentParts(Major))
    CurrentParts(Minor) = Conversion.CInt(CurrentParts(Minor))
    CurrentParts(Patch) = Conversion.CInt(CurrentParts(Patch))
    OtherParts(Major) = Conversion.CInt(OtherParts(Major))
    OtherParts(Minor) = Conversion.CInt(OtherParts(Minor))
    OtherParts(Patch) = Conversion.CInt(OtherParts(Patch))

    If CurrentParts(Major) > OtherParts(Major) Then
        GetNewest = Current
    ElseIf CurrentParts(Major) < OtherParts(Major) Then
        GetNewest = Other
    ElseIf CurrentParts(Minor) > OtherParts(Minor) Then
        GetNewest = Current
    ElseIf CurrentParts(Minor) < OtherParts(Minor) Then
        GetNewest = Other
    ElseIf CurrentParts(Patch) > OtherParts(Patch) Then
        GetNewest = Current
    ElseIf CurrentParts(Patch) < OtherParts(Patch) Then
        GetNewest = Other
    Else
        GetNewest = Current
    End If
End Function

Public Function GetLocalLibs(ByVal Registry As String, ByRef Libs As Dictionary) As Variant
    If Strings.Right(Registry, 1) <> Application.PathSeparator Then
        Registry = Registry & Application.PathSeparator
    End If

    Dim Paths As Collection: Set Paths = New Collection
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim CacheFolder As String
    CacheFolder = FSO.BuildPath(Constants.UserConfigPath, "cache")
    PFileSystem.CreateFolder CacheFolder, Recoursive:=True

    Dim Name As Variant
    For Each Name In Libs
        Dim Version As String: Version = Libs(Name)
        Dim FileName As String
        FileName = PStrings.FString("{0}_{1}.zip", Name, Strings.Replace(Version, ".", "_"))
        Dim Path As String
        Path = PStrings.FString("{0}{1}{2}{3}", Registry, Name, Application.PathSeparator, FileName)

        Dim SavePath As String
        SavePath = FSO.BuildPath(CacheFolder, FileName)
        FileSystem.FileCopy Path, SavePath
        Paths.Add SavePath
    Next

    GetLocalLibs = PCollection.ToArray(Paths)
End Function

Public Function DownloadLibs(ByVal Registry As String, ByRef Libs As Dictionary) As Variant
    If Strings.Right(Registry, 1) <> "/" Then
        Registry = Registry & "/download"
    Else
        Registry = Registry & "download"
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
        URL = PStrings.FString("{0}/{1}/{2}", Registry, Name, Version)
        Dim Response As TResponse
        Response = HTTP.GetRequest(URL)
        If Response.Code <> 200 Then
            Information.Err.Raise _
                17, _
                Information.TypeName(Me), _
                "ERR: download failed: " & PStrings.FString("{0}@{1}\\n{2}", Name, Version, Response.Text)
        End If

        Dim FileName As String
        FileName = PStrings.FString("{0}_{1}.zip", Name, Strings.Replace(Version, ".", "_"))
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
        Zip.SetZip Path
        Dim TempFolder As String
        TempFolder = FSO.BuildPath(FSO.GetParentFolderName(Path), FSO.GetTempName())
        Zip.Unzip TempFolder
        Me.ImportModules TempFolder
        FileSystem.RmDir TempFolder
        FileSystem.Kill Path
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



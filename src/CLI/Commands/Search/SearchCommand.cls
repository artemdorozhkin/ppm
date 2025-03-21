VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Search"
Option Explicit

Implements ICommand

Private Type TSearchCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TSearchCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/search"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "search"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("local") = CLI.Lang.GetValue("param", "name=local")
    this.CommandInfo.Params.Item("registry") = CLI.Lang.GetValue("param", "name=registry")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If this.Tokens.Count = 1 Then
        ppm "search -h"
        Exit Sub
    End If

    Dim Registry As String
    Registry = Utils.GetTokenOrConfigValue("registry", this.Tokens)

    Dim IsLocal As Boolean
    IsLocal = Utils.GetTokenOrConfigValue("local", this.Tokens, CastTo:=VbVarType.vbBoolean)

    Dim PackageToken As SyntaxToken
    Set PackageToken = this.Tokens.PopFirstTokenKind(TokenKind.Identifier)
    If IsFalse(PackageToken) Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/nameMissing")
        Exit Sub
    End If

    Dim PackageName As String
    PackageName = PackageToken.Text

    If IsLocal Then
        LookupLocal Registry, PackageName
    Else
        LookupOnServer Registry, PackageName
    End If
End Sub

Public Sub LookupOnServer(ByVal Registry As String, ByVal PackageName As String)
    Dim URL As String
    URL = Utils.ResolveUrl(Registry, PStrings.FString("/-/search/{0}", PackageName))

    Dim HTTP As HTTP: Set HTTP = New HTTP
    Dim Response As TResponse
    Response = HTTP.GetRequest(URL)

    If Response.Code = HTTPCodes.OK_200 Then
        Dim JSON As Object: Set JSON = PJSON.Parse(Response.Text)
        If Information.IsObject(JSON("result")) Then
            If JSON("result").Count = 0 Then
                Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/noMatchesFor"), PackageName)
            Else
                PrintResult JSON("result")
            End If
        ElseIf Information.IsNull(JSON("result")) Then
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/noMatchesFor"), PackageName)
        End If
    End If
End Sub

Public Sub LookupLocal(ByVal Registry As String, ByVal PackageName As String)
    Dim Result As Collection: Set Result = New Collection
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim RegistryFolder As Folder: Set RegistryFolder = NewFolder(Registry)
    Dim PackFolder As Folder
  #Else
    Dim FSO As Object: Set FSO = GetFileSystemObject()
    Dim RegistryFolder As Object: Set RegistryFolder = NewFolder(Registry)
    Dim PackFolder As Object
  #End If

    For Each PackFolder In RegistryFolder.SubFolders
        If PStrings.Contains(PackFolder.Name, PackageName) Then
            Dim ManifestPath As String
            ManifestPath = FSO.BuildPath(PackFolder.Path, "manifest.json")
            Dim ManifestContent As String
            ManifestContent = FSO.GetFile(ManifestPath).OpenAsTextStream().ReadAll()
            Dim Manifest As Manifest: Set Manifest = NewManifest(ManifestContent)
            If PStrings.IsEqual(PackFolder.Name, PackageName) Then
                Result.Add Manifest.JSON, Before:=1
            Else
                Result.Add Manifest.JSON
            End If
        End If
    Next

    PrintResult Result
End Sub

Public Sub PrintResult(ByRef Result As Collection)
    Dim Header As Variant: Header = Array("NAME", "VERSION", "AUTHOR", "DESCRIPTION")
    Dim Data() As Variant
    ReDim Data(0 To Result.Count - 1, 0 To UBound(Header))

    Dim Manifest As Object
    For Each Manifest In Result
        Dim i As Long
        On Error Resume Next
        Data(i, 0) = Manifest("name")
        Data(i, 1) = Manifest("latest")
        Data(i, 2) = Manifest("versions")(Manifest("latest"))("author")
        Data(i, 3) = Manifest("versions")(Manifest("latest"))("description")
        On Error GoTo 0

        i = i + 1
    Next

    Immediate.WriteTable Data, Header, MaxWidth:=20
End Sub

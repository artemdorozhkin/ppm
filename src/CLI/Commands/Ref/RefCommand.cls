VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Ref"
Option Explicit

Implements ICommand

Private Enum NameTypes
    TypeName
    TypeGUID
    TypePath
End Enum

Private Type TRefCommand
    CommandInfo As CommandInfo
    Tokens As Tokens

    Name As String
    NameType As NameTypes
End Type

Private this As TRefCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/ref"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "ref"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("save-dev") = CLI.Lang.GetValue("param", "name=save-dev")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim SaveDev As Boolean
    SaveDev = this.Tokens.IncludeDefinition(Definitions("save-dev"))

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.PopFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    End If

    If IsFalse(SubCommandName) Then
        ppm "ref --help"
        Exit Sub
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If

    If SubCommandName = "list" Then
        Me.PrintList
        Exit Sub
    ElseIf SubCommandName = "update" Then
        Me.UpdateReferences
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refsUpdated"), this.Name)
        Exit Sub
    End If

    Dim RefNames As Variant
    RefNames = this.Tokens.GetTokensKind(TokenKind.Identifier)

    If UBound(RefNames) = -1 Then
        ppm "ref --help"
        Exit Sub
    End If

    Dim Token As Variant
    For Each Token In RefNames
        this.Name = Token.Text
        If PStrings.IsEqual(this.Name, "ppm") Then
            this.Name = ThisProject.Path
        End If

        If PStrings.StartsWith(this.Name, "{") Then
            this.NameType = NameTypes.TypeGUID
        ElseIf GetFileSystemObject().FileExists(this.Name) Then
            this.NameType = NameTypes.TypePath
        ElseIf Not PStrings.Contains(this.Name, " ") Then
            this.NameType = NameTypes.TypeName
        Else
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindRef"), this.Name)
            Exit Sub
        End If

        If SubCommandName = "add" Or SubCommandName = "install" Then
            Me.AddReference
            Me.UpdateReferences SaveDev
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refAdded"), this.Name)
        ElseIf SubCommandName = "delete" Then
            Me.RmReference
            Me.UpdateReferences
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refDeleted"), this.Name)
        Else
            ppm "ref --help"
            Exit Sub
        End If
    Next
End Sub

Public Function ReadReferences() As Collection
    FileSystem.ChDir PFileSystem.BuildPath(Interaction.Environ("APPDATA"), "ppm")
    Dim RefsString As String
    RefsString = PFileSystem.ReadFile("./references/references.json")

    Set ReadReferences = PJSON.Parse(RefsString)
End Function

Public Sub AddReference()
    Dim Refs As Collection
    Set Refs = Me.ReadReferences()

    Dim Ref As Object
    On Error Resume Next
    If this.NameType = NameTypes.TypeGUID Then
        Set Ref = SelectedProject.References.AddFromGuid(this.Name, Major:=0, Minor:=0)
    ElseIf this.NameType = NameTypes.TypePath Then
        Set Ref = SelectedProject.References.AddFromFile(this.Name)
    ElseIf this.NameType = NameTypes.TypeName Then
        Dim r As Object
        For Each r In Refs
            If Strings.LCase(r("name")) = Strings.LCase(this.Name) Then
                Set Ref = SelectedProject.References.AddFromGuid(r("guid"), Major:=0, Minor:=0)
                Exit For
            End If
        Next
    End If

    If Information.Err.Number = 32813 Then
        Immediate.WriteLine "ERR: Reference already exists in the project:", this.Name
        On Error GoTo 0
        Exit Sub
    ElseIf Information.Err.Number <> 0 Then
        Dim ErrNumber As Long
        ErrNumber = Information.Err.Number
        Dim ErrSource As String
        ErrSource = Information.Err.Source
        Dim ErrDescription As String
        ErrDescription = Information.Err.Description

        On Error GoTo 0
        Information.Err.Raise ErrNumber, ErrSource, ErrDescription
    End If

    If Ref Is Nothing Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindRef"), this.Name)
        End
    End If
End Sub

Public Sub RmReference()
    Dim r As Object
    If this.NameType = NameTypes.TypeName Then
        On Error Resume Next
        Set r = SelectedProject.References(this.Name)
    ElseIf this.NameType = NameTypes.TypeGUID Then
        For Each r In SelectedProject.References
            If Strings.LCase(r.Guid) = Strings.LCase(this.Name) Then Exit For
        Next
    ElseIf this.NameType = NameTypes.TypePath Then
        For Each r In SelectedProject.References
            If Strings.LCase(r.FullPath) = Strings.LCase(this.Name) Then Exit For
        Next
    End If
    SelectedProject.References.Remove r

    If r Is Nothing Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindRef"), this.Name)
        End
    End If
End Sub

Public Sub UpdateReferences(Optional ByVal SaveDev As Boolean = False)
    If Not SelectedProject.IsComponentExists("package") Then
        SelectedProject.AddComponent "package"
    End If
    Dim Pack As Pack
    Set Pack = NewPack(SelectedProject.GetComponent("package"))

    Dim References As Object
    If IsTrue(Pack.References) Then
        Set References = Pack.References
    Else
        Set References = NewDictionary()
    End If

    Dim DevReferences As Object
    If IsTrue(Pack.DevReferences) Then
        Set DevReferences = Pack.DevReferences
    Else
        Set DevReferences = NewDictionary()
    End If

    Dim Dummy As Variant
    Dim Name As Variant
    For Each Name In References
        If Not SelectedProject.IsReferenceExists(Name) Then
            References.Remove Name
        End If
    Next

    For Each Name In DevReferences
        If Not SelectedProject.IsReferenceExists(Name) Then
            DevReferences.Remove Name
        End If
    Next

    Dim r As Object
    For Each r In SelectedProject.References
        If References.Exists(r.Name) Then GoTo Continue
        If DevReferences.Exists(r.Name) Then GoTo Continue

        If r.Name = ThisProject.Name Then
            Dim ThisPack As Pack: Set ThisPack = NewPack(ThisProject.GetComponent("package"))
            DevReferences("ppm") = ThisPack.Version
            GoTo Continue
        End If

        If SaveDev Then
            DevReferences(r.Name) = FString("{0}.{1}", r.Major, r.Minor)
        Else
            References(r.Name) = FString("{0}.{1}", r.Major, r.Minor)
        End If
Continue:
    Next

    If References.Count = 0 Then Set References = Nothing
    If DevReferences.Count = 0 Then Set DevReferences = Nothing
    Pack.Update References:=References, DevReferences:=DevReferences
End Sub

Public Sub PrintList()
    Dim Header As Variant: Header = Array("NAME", "VERSION")
    Dim Data() As Variant
    ReDim Data(0 To SelectedProject.References.Count - 1, 0 To UBound(Header))

    Dim r As Object
    For Each r In SelectedProject.References
        Dim i As Long
        On Error Resume Next
        If r.Name = ThisProject.Name Then
            Dim ThisPack As Pack: Set ThisPack = NewPack(ThisProject.GetComponent("package"))
            Data(i, 0) = "ppm"
            Data(i, 1) = ThisPack.Version
        Else
            Data(i, 0) = r.Name
            Data(i, 1) = FString("{0}.{1}", r.Major, r.Minor)
        End If
        On Error GoTo 0

        i = i + 1
    Next

    Immediate.WriteTable Data, Header, MaxWidth:=20
End Sub

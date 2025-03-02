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

Private Type TConfigCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens

    Name As String
    NameType As NameTypes
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/ref"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "ref"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
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
    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.PopFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If

    If this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        this.Name = this.Tokens.PopFirstTokenKind(TokenKind.Identifier).Text
        If PStrings.StartsWith(this.Name, "{") Then
            this.NameType = NameTypes.TypeGUID
        ElseIf NewFileSystemObject().FileExists(this.Name) Then
            this.NameType = NameTypes.TypePath
        ElseIf Not PStrings.Contains(this.Name, " ") Then
            this.NameType = NameTypes.TypeName
        Else
            ' TODO: raise error
        End If
    End If

    If IsFalse(SubCommandName) Or SubCommandName = "add" Then
        If Strings.Len(this.Name) = 0 Then
            ppm "ref --help"
        Else
            Me.AddReference
            Me.UpdateReferences
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refAdded"), this.Name)
        End If
        Exit Sub
    End If

    If Not this.Tokens.IncludeTokenKind(TokenKind.Identifier) And IsTrue(SubCommandName) Then
        If SubCommandName = "list" Then
            Me.PrintList
            Exit Sub
        ElseIf SubCommandName = "update" Then
            Me.UpdateReferences
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refsUpdated"), this.Name)
            Exit Sub
        Else
            ppm "ref --help"
        End If
        Exit Sub
    End If

    If IsTrue(SubCommandName) And SubCommandName = "delete" Then
        Me.RmReference
        Me.UpdateReferences
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/refDeleted"), this.Name)
    Else
        ppm "ref -- help"
    End If
End Sub

Public Function ReadReferences() As Collection
    FileSystem.ChDir SelectedProject.Folder
    Dim RefsString As String
    RefsString = PFileSystem.ReadFile("./references/references.json")

    Set ReadReferences = PJSON.Parse(RefsString)
End Function

Public Sub AddReference()
    Dim Refs As Collection
    Set Refs = Me.ReadReferences()

    Dim Ref As Object
    If this.NameType = NameTypes.TypeGUID Then
        Set Ref = ThisWorkbook.VBProject.References.AddFromGuid(this.Name, Major:=0, Minor:=0)
    ElseIf this.NameType = NameTypes.TypePath Then
        Set Ref = ThisWorkbook.VBProject.References.AddFromFile(this.Name)
    ElseIf this.NameType = NameTypes.TypeName Then
        Dim r As Object
        For Each r In Refs
            If Strings.LCase(r("name")) = Strings.LCase(this.Name) Then
                Set Ref = ThisWorkbook.VBProject.References.AddFromGuid(r("guid"), Major:=0, Minor:=0)
                Exit For
            End If
        Next
    End If

    If Not Ref Is Nothing Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindRef"), this.Name)
        End
    End If
End Sub

Public Sub RmReference()
    Dim r As Object
    If this.NameType = NameTypes.TypeName Then
        On Error Resume Next
        Set r = ThisWorkbook.VBProject.References(this.Name)
        ThisWorkbook.VBProject.References.Remove r
    ElseIf this.NameType = NameTypes.TypeGUID Then
        For Each r In ThisWorkbook.VBProject.References
            If Strings.LCase(r.Guid) = this.Name Then Exit For
        Next
    ElseIf this.NameType = NameTypes.TypePath Then
        For Each r In ThisWorkbook.VBProject.References
            If Strings.LCase(r.FullPath) = this.Name Then Exit For
        Next
    End If

    If r Is Nothing Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindRef"), this.Name)
        End
    End If
End Sub

Public Sub UpdateReferences()
    Dim References As Object
    Set References = NewDictionary()

    Dim r As Object
    For Each r In ThisWorkbook.VBProject.References
        References(r.Name) = FString("{0}.{1}", r.Major, r.Minor)
    Next

    Dim Pack As Pack
    Set Pack = NewPack(SelectedProject.GetComponent("package"))

    Pack.Update References:=References
End Sub

Public Sub PrintList()
    Dim Header As Variant: Header = Array("NAME", "VERSION")
    Dim Data() As Variant
    ReDim Data(0 To ThisWorkbook.VBProject.References.Count - 1, 0 To UBound(Header))

    Dim r As Object
    For Each r In ThisWorkbook.VBProject.References
        Dim i As Long
        On Error Resume Next
        Data(i, 0) = r.Name
        Data(i, 1) = FString("{0}.{1}", r.Major, r.Minor)
        On Error GoTo 0

        i = i + 1
    Next

    Immediate.WriteTable Data, Header, MaxWidth:=20
End Sub

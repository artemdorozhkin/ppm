VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Pack"
Option Explicit

Private Type TPack
    Project As Project
    Pack As Object
    Module As VBComponent
End Type

Private this As TPack

Public Property Get Project() As Project
    Set Project = this.Project
End Property
Public Property Set Project(ByVal RHS As Project)
    Set this.Project = RHS
End Property

Public Property Get JSON() As Object
    Set JSON = this.Pack
End Property

Public Property Get Module() As VBComponent
    Set Module = this.Module
End Property

Public Property Get Version() As String
    Version = this.Pack("version")
End Property

Public Property Get Name() As String
    Name = this.Pack("name")
End Property

Public Property Get Description() As String
    Description = this.Pack("description")
End Property

Public Property Get Author() As String
    Author = this.Pack("author")
End Property

Public Property Get Git() As String
    Git = this.Pack("git")
End Property

Public Property Get Dependencies() As Dictionary
    Set Dependencies = this.Pack("dependencies")
End Property

Public Property Get DevDependencies() As Dictionary
    Set DevDependencies = this.Pack("devDependencies")
End Property

Public Function Read() As Pack
    Set this.Module = this.Project.GetModule("package")
    Dim Code As String
    Code = this.Module.CodeModule.Lines(1, this.Module.CodeModule.CountOfLines)
    Dim Lines As Collection: Set Lines = New Collection
    Dim Line As Variant
    For Each Line In Strings.Split(Code, vbNewLine)
        If PStrings.StartsWith(Line, "'") And Not PStrings.Contains(Line, "@Folder") Then
            Lines.Add Line
        End If
    Next
    Code = Strings.Join(PCollection.ToArray(Lines), vbNewLine)
    Set this.Pack = PJSON.Parse(Utils.UncommentString(Code))
    Set Read = Me
End Function

Public Function RenameModule(ByVal Name As String) As Pack
    this.Module.Name = Name
    Set RenameModule = Me
End Function

Public Function Update( _
    Optional ByVal Name As Variant, _
    Optional ByVal Version As Variant, _
    Optional ByVal Description As Variant, _
    Optional ByVal Author As Variant, _
    Optional ByVal Git As Variant, _
    Optional ByVal Dependencies As Variant, _
    Optional ByVal DevDependencies As Variant _
) As Pack
    If Not Information.IsMissing(Name) Then this.Pack("name") = Name
    If Not Information.IsMissing(Version) Then this.Pack("version") = Version
    If Not Information.IsMissing(Description) Then this.Pack("description") = Description
    If Not Information.IsMissing(Author) Then this.Pack("author") = Author
    If Not Information.IsMissing(Git) Then this.Pack("git") = Git
    If Not Information.IsMissing(Dependencies) Then Set this.Pack("dependencies") = Dependencies
    If Not Information.IsMissing(DevDependencies) Then Set this.Pack("devDependencies") = DevDependencies

    Dim JSON As String: JSON = PJSON.Stringify(this.Pack, 2)
    With this.Module.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString Utils.CommentString(JSON)
    End With
    Set Update = Me
End Function



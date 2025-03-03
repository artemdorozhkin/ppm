VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, ProcedureNotUsed
'@Folder "PearPMProject.src.Pack"
Option Explicit

Private Type TPack
  #If DEV Then
    Pack As Dictionary
    Component As VBComponent
  #Else
    Pack As Object
    Component As Object
  #End If
End Type

Private this As TPack

#If DEV Then
  Public Property Get JSON() As Dictionary
#Else
  Public Property Get JSON() As Object
#End If
    Set JSON = this.Pack
End Property
#If DEV Then
  Public Property Set JSON(ByVal RHS As Dictionary)
#Else
  Public Property Set JSON(ByVal RHS As Object)
#End If
    Set this.Pack = RHS
End Property

#If DEV Then
  Public Property Get Component() As VBComponent
#Else
  Public Property Get Component() As Object
#End If
    Set Component = this.Component
End Property

Public Property Get Root() As String
    Root = NewModule(this.Component).Path
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

#If DEV Then
  Public Property Get Dependencies() As Dictionary
#Else
  Public Property Get Dependencies() As Object
#End If
    If this.Pack.Exists("dependencies") Then
        Set Dependencies = this.Pack("dependencies")
    End If
End Property

#If DEV Then
  Public Property Get DevDependencies() As Dictionary
#Else
  Public Property Get DevDependencies() As Object
#End If
    If this.Pack.Exists("devDependencies") Then
        Set DevDependencies = this.Pack("devDependencies")
    End If
End Property

#If DEV Then
  Public Property Get References() As Dictionary
#Else
  Public Property Get References() As Object
#End If
    If this.Pack.Exists("references") Then
        Set References = this.Pack("references")
    End If
End Property

#If DEV Then
  Public Function Read(ByRef Component As VBComponent) As Pack
#Else
  Public Function Read(ByRef Component As Object) As Pack
#End If
    Set this.Component = Component
    Dim Code As String
    Code = this.Component.CodeModule.Lines(1, this.Component.CodeModule.CountOfLines)

    Dim OpenedBraces As Long
    Dim Lines As Collection: Set Lines = New Collection
    Dim Line As Variant
    For Each Line In Strings.Split(Code, vbNewLine)
        Dim LineWithoutSpaces As String: LineWithoutSpaces = Strings.Replace(Line, " ", "")
        If PStrings.StartsWith(LineWithoutSpaces, "'{") Or OpenedBraces > 0 Then
            If PStrings.Contains(LineWithoutSpaces, "{") Then OpenedBraces = OpenedBraces + 1
            If PStrings.Contains(LineWithoutSpaces, "}") Then OpenedBraces = OpenedBraces - 1
            Lines.Add Line
        End If
    Next
    Code = Strings.Join(PCollection.ToArray(Lines), vbNewLine)
    Set this.Pack = PJSON.Parse(Utils.UncommentString(Code))
    If this.Pack.Exists("dependencies") Then
        Set this.Pack("dependencies") = Utils.ChangeDictionaryCompareMode( _
            this.Pack("dependencies"), _
            VbCompareMethod.vbTextCompare _
        )
    End If
    If this.Pack.Exists("devDependencies") Then
        Set this.Pack("devDependencies") = Utils.ChangeDictionaryCompareMode( _
            this.Pack("devDependencies"), _
            VbCompareMethod.vbTextCompare _
        )
    End If
    If this.Pack.Exists("references") Then
        Set this.Pack("references") = Utils.ChangeDictionaryCompareMode( _
            this.Pack("references"), _
            VbCompareMethod.vbTextCompare _
        )
    End If

    Set Read = Me
End Function

Public Function RenameComponent(ByVal Name As String) As Pack
    this.Component.Name = Name
    Set RenameComponent = Me
End Function

Public Function Update( _
    Optional ByVal Name As Variant, _
    Optional ByVal Version As Variant, _
    Optional ByVal Description As Variant, _
    Optional ByVal Author As Variant, _
    Optional ByVal Git As Variant, _
    Optional ByVal Dependencies As Variant, _
    Optional ByVal DevDependencies As Variant, _
    Optional ByVal References As Variant _
) As Pack
    If Not Information.IsMissing(Name) Then this.Pack("name") = Name
    If Not Information.IsMissing(Version) Then this.Pack("version") = Version
    If Not Information.IsMissing(Description) Then this.Pack("description") = Description
    If Not Information.IsMissing(Author) Then this.Pack("author") = Author
    If Not Information.IsMissing(Git) Then this.Pack("git") = Git

    Dim DepName As Variant
    If Not Information.IsMissing(Dependencies) Then
        If this.Pack.Exists("dependencies") Then
            
            For Each DepName In Dependencies
                this.Pack("dependencies")(DepName) = Dependencies(DepName)
            Next
        ElseIf Not Dependencies Is Nothing Then
            Set this.Pack("dependencies") = Dependencies
        End If
    End If

    If Not Information.IsMissing(DevDependencies) Then
        If this.Pack.Exists("devDependencies") Then
            For Each DepName In DevDependencies
                this.Pack("devDependencies")(DepName) = DevDependencies(DepName)
            Next
        ElseIf Not DevDependencies Is Nothing Then
            Set this.Pack("devDependencies") = DevDependencies
        End If
    End If

    If Not Information.IsMissing(References) Then
        If Not References Is Nothing Then
            Set this.Pack("references") = References
        End If
    End If

    Dim JSON As String: JSON = PJSON.Stringify(this.Pack, 2)
    With this.Component.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString PStrings.FString("'@Folder(""{0}"")", SelectedProject.Name)
        .AddFromString Utils.CommentString(JSON)
    End With
    Set Update = Me
End Function

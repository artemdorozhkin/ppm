VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Project"
Option Explicit

Private Type TProject
  #If DEV Then
    Project As VBProject
  #Else
    Project As Object
  #End If
    FileName As String
End Type

Private this As TProject

#If DEV Then
  Public Property Get Components() As VBComponents
#Else
  Public Property Get Components() As Object
#End If
    Set Components = this.Project.VBComponents
End Property

#If DEV Then
  Public Property Get References() As References
#Else
  Public Property Get References() As Object
#End If
    Set References = this.Project.References
End Property

Public Property Get IsSaved() As Boolean
    On Error Resume Next
    IsSaved = this.FileName <> ""
End Property

Public Property Get Name() As String
    Name = this.Project.Name
End Property

Public Property Get Path() As String
    Path = this.FileName
End Property

Public Property Get Folder() As String
    Folder = NewFolder(this.FileName).Path
End Property

Public Sub Rename(ByVal Name As String)
    this.Project.Name = Name
End Sub

#If DEV Then
  Public Property Set Project(ByVal RHS As VBProject)
#Else
  Public Property Set Project(ByVal RHS As Object)
#End If
    Set this.Project = RHS
    On Error Resume Next
    this.FileName = this.Project.FileName
End Property

#If DEV Then
  Public Function GetComponent(ByVal Name As String) As VBComponent
#Else
  Public Function GetComponent(ByVal Name As String) As Object
#End If
    If Not Me.IsComponentExists(Name) Then Exit Function
    Set GetComponent = this.Project.VBComponents(Name)
End Function

#If DEV Then
  Public Function AddComponent( _
      ByVal Name As String, _
      Optional ByVal Kind As ModuleType = ModuleType.StdModule _
  ) As VBComponent
#Else
  Public Function AddComponent( _
      ByVal Name As String, _
      Optional ByVal Kind As ModuleType = ModuleType.StdModule _
  ) As Object
#End If
    Set AddComponent = this.Project.VBComponents.Add(Kind)
    AddComponent.Name = Name
End Function

#If DEV Then
  Public Function ImportComponent(ByVal Path As String) As VBComponent
#Else
  Public Function ImportComponent(ByVal Path As String) As Object
#End If
    If Not IsValidType(Path) Then
        Information.Err.Raise _
            13, _
            Information.TypeName(Me), _
            "unsupproted type: '" & GetFileExt(Path) & "'"
    End If
    If IsComponentExists(Path) Then
        Information.Err.Raise _
            58, _
            Information.TypeName(Me), _
            "this module already exists: '" & GetFileName(Path) & "'"
    End If
    Set ImportComponent = this.Project.VBComponents.Import(Path)
End Function

Public Sub RemoveComponent(ByVal Name As String)
    If Not IsComponentExists(Name) Then Exit Sub
    this.Project.VBComponents.Remove this.Project.VBComponents(Name)
End Sub

#If DEV Then
  Public Function ExportComponent(ByVal Folder As String, ByRef Component As VBComponent) As String
#Else
  Public Function ExportComponent(ByVal Folder As String, ByRef Component As Object) As String
#End If
    Dim Ext As String: Ext = NewModule(Component).Ext
    Dim Path As String: Path = GetFileSystemObject().BuildPath(Folder, Component.Name & Ext)
    TrimCodeEnd Component
    Component.Export Path
    ExportComponent = Path
End Function

Public Function IsComponentExists(ByVal NameOrPath As String) As Boolean
    On Error GoTo Catch
    Dim Name As String: Name = GetFileNameWithoutExt(NameOrPath)
    IsComponentExists = this.Project.VBComponents(Name).Name <> ""
Exit Function

Catch:
    IsComponentExists = False
End Function

Public Function IsReferenceExists(ByVal Name As String) As Boolean
    On Error GoTo Catch
    IsReferenceExists = this.Project.References(Name).Name <> ""
Exit Function

Catch:
    IsReferenceExists = False
End Function

Private Function IsValidType(ByVal Path As String) As Boolean
    Dim Ext As String: Ext = Strings.LCase(GetFileExt(Path))
    IsValidType = Ext = ".bas" Or Ext = ".cls" Or Ext = ".frm"
End Function

#If DEV Then
  Private Sub TrimCodeEnd(ByRef Component As VBComponent)
#Else
  Private Sub TrimCodeEnd(ByRef Component As Object)
#End If
    Dim Module As Module: Set Module = NewModule(Component)
    Dim Lines As Variant: Lines = Strings.Split(Module.Code, vbNewLine)

    Dim Count As Long: Count = UBound(Lines)
    Dim i As Long
    For i = Count To 0 Step -1
        Dim TrimmedLine As String
        TrimmedLine = Strings.Trim(Lines(i))
        If Strings.Len(TrimmedLine) = 0 Then GoTo Continue
        i = i + 1
        Exit For
Continue:
    Next

    If Count < i Then Exit Sub

    If i = -1 Then
        Module.Code = ""
    Else
        Component.CodeModule.DeleteLines i + 1, Count + 1 - i
    End If
End Sub

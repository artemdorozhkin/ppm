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
    
Public Enum ModuleType
    StdModule = 1
    ClassModule
    FormModule
    DocModule = 100
End Enum

Private Type TProject
  #If DEV Then
    Project As VBProject
  #Else
    Project As Object
  #End If
End Type

Private this As TProject

#If DEV Then
  Public Property Get Components() As VBComponents
#Else
  Public Property Get Components() As Object
#End If
    Set Components = this.Project.VBComponents
End Property

Public Property Get Name() As String
    Name = this.Project.Name
End Property

Public Property Get Path() As String
    Path = this.Project.FileName
End Property

Public Property Get Folder() As String
    Folder = NewFolder(this.Project.FileName).Path
End Property

#If DEV Then
  Public Property Set Project(ByVal RHS As VBProject)
#Else
  Public Property Set Project(ByVal RHS As Object)
#End If
    Set this.Project = RHS
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
  Public Function AddComponent(ByVal Name As String) As VBComponent
#Else
  Public Function AddComponent(ByVal Name As String) As Object
#End If
    Set AddComponent = this.Project.VBComponents.Add(ModuleType.StdModule)
    AddComponent.Name = Name
End Function

#If DEV Then
  Public Function ImportComponent(ByVal Path As String) As VBComponent
#Else
  Public Function ImportComponent(ByVal Path As String) As Object
#End If
    If Not IsValidType(Path) Then Information.Err.Raise 13, Information.TypeName(Me), "unsupproted type: '" & GetFileExt(Path) & "'"
    If IsComponentExists(Path) Then Information.Err.Raise 58, Information.TypeName(Me), "this module already exists: '" & GetFileName(Path) & "'"
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
    Dim Ext As String: Ext = GetComponentExt(Component)
    Dim Path As String: Path = NewFileSystemObject().BuildPath(Folder, Component.Name & Ext)
    Component.Export Path
    ExportComponent = Path
End Function

Public Function IsComponentExists(ByVal NameOrPath As String) As Boolean
    On Error GoTo Catch
    Dim Name As String: Name = GetFileNameWithoutExt(NameOrPath)
    IsComponentExists = this.Project.VBComponents(Name).Name <> ""
Exit Function

Catch:
    If Information.Err.Number = 9 Then
        IsComponentExists = False
    Else
        Information.Err.Raise Information.Err.Number, Information.TypeName(Me)
    End If
End Function

Private Function IsValidType(ByVal Path As String) As Boolean
    Dim Ext As String: Ext = Strings.LCase(GetFileExt(Path))
    IsValidType = Ext = ".bas" Or Ext = ".cls"
End Function

#If DEV Then
  Private Function GetComponentExt(ByRef Module As VBComponent) As String
#Else
  Private Function GetComponentExt(ByRef Module As Object) As String
#End If
    Select Case Module.Type
    Case ModuleType.StdModule: GetComponentExt = ".bas"
    Case ModuleType.ClassModule: GetComponentExt = ".cls"
    Case ModuleType.FormModule: GetComponentExt = ".frm"
    Case ModuleType.DocModule: GetComponentExt = ".doccls"
    Case Else: GetComponentExt = ".txt"
    End Select
End Function

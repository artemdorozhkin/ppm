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
  Public Property Get Modules() As VBComponents
#Else
  Public Property Get Modules() As Object
#End If
    Set Modules = this.Project.VBComponents
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

Public Property Set Project(ByVal RHS As VBProject)
    Set this.Project = RHS
End Property

#If DEV Then
  Public Function GetModule(ByVal Name As String) As VBComponent
#Else
  Public Function GetModule(ByVal Name As String) As Object
#End If
    If Not Me.IsModuleExists(Name) Then Exit Function
    Set GetModule = this.Project.VBComponents(Name)
End Function

#If DEV Then
  Public Function AddModule(ByVal Name As String) As VBComponent
#Else
  Public Function AddModule(ByVal Name As String) As Object
#End If
    Set AddModule = this.Project.VBComponents.Add(ModuleType.StdModule)
    AddModule.Name = Name
End Function

Public Function ImportModule(ByVal Path As String)
    If Not IsValidType(Path) Then Err.Raise 13, TypeName(Me), "unsupproted type: '" & GetFileExt(Path) & "'"
    If IsModuleExists(Path) Then Err.Raise 58, TypeName(Me), "this module already exists: '" & GetFileName(Path) & "'"
    Set ImportModule = this.Project.VBComponents.Import(Path)
End Function

#If DEV Then
  Public Function ExportModule(ByVal Folder As String, ByRef Module As VBComponent) As String
#Else
  Public Function ExportModule(ByVal Folder As String, ByRef Module As Object) As String
#End If
    Dim Ext As String: Ext = GetModuleExt(Module)
    Dim Path As String: Path = NewFileSystemObject().BuildPath(Folder, Module.Name & Ext)
    Module.Export Path
    ExportModule = Path
End Function

Public Function IsModuleExists(ByVal NameOrPath As String) As Boolean
    On Error GoTo Catch
    Dim Name As String: Name = GetFileNameWithoutExt(NameOrPath)
    IsModuleExists = this.Project.VBComponents(Name).Name <> ""
Exit Function

Catch:
    If Err.Number = 9 Then
        IsModuleExists = False
    Else
        Err.Raise Err.Number, TypeName(Me)
    End If
End Function

Private Function IsValidType(ByVal Path As String) As Boolean
    Dim Ext As String: Ext = Strings.LCase(GetFileExt(Path))
    IsValidType = Ext = ".bas" Or Ext = ".cls"
End Function

#If DEV Then
  Private Function GetModuleExt(ByRef Module As VBComponent) As String
#Else
  Private Function GetModuleExt(ByRef Module As Object) As String
#End If
    Select Case Module.Type
    Case ModuleType.StdModule: GetModuleExt = ".bas"
    Case ModuleType.ClassModule: GetModuleExt = ".cls"
    Case ModuleType.FormModule: GetModuleExt = ".frm"
    Case ModuleType.DocModule: GetModuleExt = ".doccls"
    Case Else: GetModuleExt = ".txt"
    End Select
End Function

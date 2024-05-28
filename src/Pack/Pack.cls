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

Public Property Get Module() As VBComponent
    Set Module = this.Module
End Property

Public Property Get Version() As String
    Version = this.Pack("version")
End Property

Public Property Get Name() As String
    Version = this.Pack("name")
End Property

Public Property Get Description() As String
    Version = this.Pack("description")
End Property

Public Property Get Author() As String
    Version = this.Pack("author")
End Property

Public Property Get Git() As String
    Version = this.Pack("git")
End Property

Public Sub Read()
    Set this.Module = this.Project.GetModule("package")
    Dim Code As String
    Code = this.Module.CodeModule.Lines(1, this.Module.CodeModule.CountOfLines)
    Set this.Pack = PJSON.Parse(Utils.UncommentString(Code))
End Sub

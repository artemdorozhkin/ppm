VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Manifest"
Option Explicit

Private Type TManifest
    JSON As Object
End Type

Private this As TManifest

Public Property Get JSON() As Object
    Set JSON = this.JSON
End Property
Public Property Set JSON(ByVal RHS As Object)
    Set this.JSON = RHS
End Property

Public Property Get Name() As String
    Name = this.JSON("name")
End Property
Public Property Let Name(ByVal RHS As String)
    this.JSON("name") = RHS
End Property

Public Property Get Latest() As String
    Latest = this.JSON("latest")
End Property
Public Property Let Latest(ByVal RHS As String)
    this.JSON("latest") = RHS
End Property

Public Property Get Versions() As Object
    Set Versions = this.JSON("versions")
End Property
Public Property Set Versions(ByVal RHS As Object)
    Set this.JSON("versions") = RHS
End Property

Public Function ToString(Optional ByVal Indent As String = 2) As String
    ToString = PJSON.Stringify(this.JSON, Indent:=Indent)
End Function

Public Function FromString(ByVal JSONString As String) As Manifest
    Set this.JSON = PJSON.Parse(JSONString)
    Set FromString = Me
End Function

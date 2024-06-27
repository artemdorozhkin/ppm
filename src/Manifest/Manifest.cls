VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.Manifest"
Option Explicit

Private Type TManifest
  #If DEV Then
    JSON As Dictionary
  #Else
    JSON As Object
  #End If
End Type

Private this As TManifest

#If DEV Then
  Public Property Get JSON() As Dictionary
#Else
  Public Property Get JSON() As Object
#End If
    Set JSON = this.JSON
End Property
#If DEV Then
  Public Property Set JSON(ByVal RHS As Dictionary)
#Else
  Public Property Set JSON(ByVal RHS As Object)
#End If
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

#If DEV Then
  Public Property Get Versions() As Dictionary
#Else
  Public Property Get Versions() As Object
#End If
    Set Versions = this.JSON("versions")
End Property
#If DEV Then
  Public Property Set Versions(ByVal RHS As Dictionary)
#Else
  Public Property Set Versions(ByVal RHS As Object)
#End If
    Set this.JSON("versions") = RHS
End Property

Public Function ToString(Optional ByVal Indent As String = 2) As String
    ToString = PJSON.Stringify(this.JSON, Indent:=Indent)
End Function

Public Function FromString(ByVal JSONString As String) As Manifest
    Set this.JSON = PJSON.Parse(JSONString)
    Set FromString = Me
End Function

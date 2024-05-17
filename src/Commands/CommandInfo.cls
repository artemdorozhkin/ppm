VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Commands"
Option Explicit

Public Name As String
Public Description As String
Public Params As Variant
Public Usage As Variant

Private Sub Class_Initialize()
    Set Me.Params = NewDictionary()
End Sub

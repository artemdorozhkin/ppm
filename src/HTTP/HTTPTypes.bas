Attribute VB_Name = "HTTPTypes"
'@Folder "PearPMProject.src.HTTP"
Option Explicit

Public Type TResponse
    Body() As Byte
    Text As String
    Code As Long
End Type

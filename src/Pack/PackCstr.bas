Attribute VB_Name = "PackCstr"
'@Folder "PearPMProject.src.Pack"
Option Explicit

Public Function NewPack(ByRef JSONOrModule As Object) As Pack
    Set NewPack = New Pack
    If Utils.IsVBComponent(JSONOrModule) Then
        NewPack.Read JSONOrModule
    ElseIf Utils.IsDictionary(JSONOrModule) Then
        Set NewPack.JSON = JSONOrModule
    End If
End Function

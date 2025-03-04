Attribute VB_Name = "PackCstr"
'@Folder "PearPMProject.src.Pack"
Option Explicit

Public Function NewPack(ByRef JSONOrComponent As Object) As Pack
    Set NewPack = New Pack
    If Utils.IsVBComponent(JSONOrComponent) Then
        NewPack.Read JSONOrComponent
    ElseIf Utils.IsDictionary(JSONOrComponent) Then
        Set NewPack.JSON = JSONOrComponent
    Else
        Information.Err.Raise 91, "PackCstr.NewPack", "package component not set"
    End If
End Function

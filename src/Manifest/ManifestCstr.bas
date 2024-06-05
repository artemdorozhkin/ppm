Attribute VB_Name = "ManifestCstr"
'@Folder "PearPMProject.src.Manifest"
Option Explicit

Public Function NewManifest(Optional ByVal JSONObjectOrString As Variant) As Manifest
    Set NewManifest = New Manifest
    If Information.IsMissing(JSONObjectOrString) Then Exit Function

    If Information.IsObject(JSONObjectOrString) Then
        Set NewManifest.JSON = JSONObjectOrString
    ElseIf Information.VarType(JSONObjectOrString) = VbVarType.vbString Then
        NewManifest.FromString JSONObjectOrString
    End If
End Function

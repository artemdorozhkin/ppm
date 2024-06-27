Attribute VB_Name = "DefinitionCstr"
'@Folder "PearPMProject.src.Config.Definition"
Option Explicit

Public Function NewDefinition( _
    ByVal Key As String, _
    Optional ByVal Short As String, _
    Optional ByVal DataType As VbVarType = VbVarType.vbEmpty, _
    Optional ByVal Default As Variant, _
    Optional ByVal Description As String _
) As Definition
    If DataType = VbVarType.vbEmpty And _
       Not Information.IsMissing(Default) Then
       DataType = Information.VarType(Default)
    ElseIf DataType = VbVarType.vbEmpty Then
        DataType = VbVarType.vbVariant
    End If

    Set NewDefinition = New Definition
    With NewDefinition
        .Key = Key
        .Short = Short
        .DataType = DataType
        .Default = Default
        .Description = Description
    End With
End Function

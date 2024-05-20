Attribute VB_Name = "DefinitionCstr"
'@Folder "PearPMProject.src.CLI.Definition"
Option Explicit

Public Function NewDefinition( _
    ByVal Key As String, _
    ByVal KeyType As VbVarType, _
    ByVal Short As String, _
    ByVal Default As Variant, _
    Optional ByVal Description As String _
) As Definition
    Set NewDefinition = New Definition
    With NewDefinition
        .Key = Key
        .KeyType = KeyType
        .Short = Short
        .Default = Default
        .Description = Description
    End With
End Function



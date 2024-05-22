Attribute VB_Name = "ConfigCstr"
'@Folder "PearPMProject.src.Config"
Option Explicit

Public Function NewConfig(ByRef Definitions As Object, Optional ByRef Tokens As Tokens) As Config
    Set NewConfig = New Config
    Set NewConfig.Definitions = Definitions
    Set NewConfig.Tokens = Tokens
End Function

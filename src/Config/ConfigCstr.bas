Attribute VB_Name = "ConfigCstr"
'@Folder "PearPMProject.src.Config"
Option Explicit

#If DEV Then
  Public Function NewConfig(ByRef Definitions As Dictionary, Optional ByRef Tokens As Tokens) As Config
#Else
  Public Function NewConfig(ByRef Definitions As Object, Optional ByRef Tokens As Tokens) As Config
#End If
    Set NewConfig = New Config
    Set NewConfig.Definitions = Definitions
    Set NewConfig.Tokens = Tokens
End Function

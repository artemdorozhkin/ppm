Attribute VB_Name = "ConfigCstr"
'@Folder "PearPMProject.src.Config"
Option Explicit

#If DEV Then
  Public Function NewConfig(ByVal Scope As ConfigScopes) As Config
#Else
  Public Function NewConfig(ByVal Scope As ConfigScopes) As Config
#End If
    Set NewConfig = New Config
    NewConfig.SetScope Scope
End Function

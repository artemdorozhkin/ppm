Attribute VB_Name = "Configs"
'@Folder "PearPMProject.src.Config"
Option Explicit

Public Function GetProjectConfig() As ConfigIO
    Set GetProjectConfig = NewConfigIO(Constants.ProjectConfigPath)
End Function

Public Function GetUserConfig() As ConfigIO
    Set GetUserConfig = NewConfigIO(Constants.UserConfigPath)
End Function

Public Function GetGlobalConfig() As ConfigIO
    Set GetGlobalConfig = NewConfigIO(Constants.GlobalConfigPath)
End Function


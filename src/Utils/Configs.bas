Attribute VB_Name = "Configs"
'@Folder("PearPMProject.src.Utils")
Option Explicit

Public Function GetCurrentProjectConfig() As ConfigIO
    Set GetCurrentProjectConfig = NewConfigIO(PPMPaths.GetCurrentProjectConfigPath())
End Function

Public Function GetGlobalConfig() As ConfigIO
    Set GetGlobalConfig = NewConfigIO(PPMPaths.GetGlobalConfigPath())
End Function


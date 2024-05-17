Attribute VB_Name = "ConfigIOCstr"
'@Folder "PearPMProject.src.ConfigIO"
Option Explicit

Public Function NewConfigIO(ByVal Path As String) As ConfigIO
    Set NewConfigIO = New ConfigIO
    NewConfigIO.ConfigPath = Path
End Function

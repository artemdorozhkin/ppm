Attribute VB_Name = "ModuleCstr"
'@Folder "PearPMProject.src.Module"
Option Explicit

Public Function NewModule(ByRef Component As VBComponent) As Module
    Set NewModule = New Module
    Set NewModule.Item = Component
End Function

Attribute VB_Name = "ModuleCstr"
'@Folder "PearPMProject.src.Module"
Option Explicit

#If DEV Then
  Public Function NewModule(ByRef Component As VBComponent) As Module
#Else
  Public Function NewModule(ByRef Component As Object) As Module
#End If
    Set NewModule = New Module
    Set NewModule.Item = Component
End Function

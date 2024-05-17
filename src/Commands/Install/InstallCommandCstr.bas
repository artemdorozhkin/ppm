Attribute VB_Name = "InstallCommandCstr"
'@Folder "PearPMProject.src.Commands.Install"
Option Explicit

Public Function NewInstallCommand(ByRef Args As Variant) As InstallCommand
    Set NewInstallCommand = New InstallCommand
    NewInstallCommand.Args = Args
End Function

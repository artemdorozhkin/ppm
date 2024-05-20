Attribute VB_Name = "InstallCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Install"
Option Explicit

Public Function NewInstallCommand(ByRef Tokens As Variant) As InstallCommand
    Set NewInstallCommand = New InstallCommand
    Set NewInstallCommand.Tokens = Tokens
End Function

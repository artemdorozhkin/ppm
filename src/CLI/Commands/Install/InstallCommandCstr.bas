Attribute VB_Name = "InstallCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Install"
Option Explicit

Public Function NewInstallCommand(ByRef Config As Config, ByRef Tokens As Tokens) As InstallCommand
    Set NewInstallCommand = New InstallCommand
    Set NewInstallCommand.Config = Config
    Set NewInstallCommand.Tokens = Tokens
End Function

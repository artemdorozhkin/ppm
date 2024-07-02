Attribute VB_Name = "UninstallCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Uninstall"
Option Explicit

Public Function NewUninstallCommand(ByRef Config As Config, ByRef Tokens As Tokens) As UninstallCommand
    Set NewUninstallCommand = New UninstallCommand
    Set NewUninstallCommand.Config = Config
    Set NewUninstallCommand.Tokens = Tokens
End Function

Attribute VB_Name = "UninstallCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Uninstall"
Option Explicit

Public Function NewUninstallCommand(ByRef Tokens As Tokens) As UninstallCommand
    Set NewUninstallCommand = New UninstallCommand
    Set NewUninstallCommand.Tokens = Tokens
End Function

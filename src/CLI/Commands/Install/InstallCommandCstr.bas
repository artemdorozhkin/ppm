Attribute VB_Name = "InstallCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Install"
Option Explicit

Public Function NewInstallCommand(ByRef Tokens As Tokens) As InstallCommand
    Set NewInstallCommand = New InstallCommand
    Set NewInstallCommand.Tokens = Tokens
End Function

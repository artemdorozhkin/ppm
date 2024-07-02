Attribute VB_Name = "VersionCommandCstr"
'@Folder("PearPMProject.src.CLI.Commands.Version")
Option Explicit

Public Function NewVersionCommand(ByRef Config As Config, ByRef Tokens As Tokens) As VersionCommand
    Set NewVersionCommand = New VersionCommand
    Set NewVersionCommand.Config = Config
    Set NewVersionCommand.Tokens = Tokens
End Function

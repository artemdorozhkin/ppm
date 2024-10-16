Attribute VB_Name = "ClassCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Class"
Option Explicit

Public Function NewClassCommand(ByRef Config As Config, ByRef Tokens As Tokens) As ClassCommand
    Set NewClassCommand = New ClassCommand
    Set NewClassCommand.Config = Config
    Set NewClassCommand.Tokens = Tokens
End Function

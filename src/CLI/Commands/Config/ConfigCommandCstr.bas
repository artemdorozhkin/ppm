Attribute VB_Name = "ConfigCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Config"
Option Explicit

Public Function NewConfigCommand(ByRef Config As Config, ByRef Tokens As Tokens) As ConfigCommand
    Set NewConfigCommand = New ConfigCommand
    Set NewConfigCommand.Config = Config
    Set NewConfigCommand.Tokens = Tokens
End Function

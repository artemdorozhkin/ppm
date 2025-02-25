Attribute VB_Name = "RefCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Ref"
Option Explicit

Public Function NewRefCommand(ByRef Config As Config, ByRef Tokens As Tokens) As RefCommand
    Set NewRefCommand = New RefCommand
    Set NewRefCommand.Tokens = Tokens
    Set NewRefCommand.Config = Config
End Function

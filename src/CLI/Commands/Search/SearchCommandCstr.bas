Attribute VB_Name = "SearchCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Search"
Option Explicit

Public Function NewSearchCommand(ByRef Config As Config, ByRef Tokens As Tokens) As SearchCommand
    Set NewSearchCommand = New SearchCommand
    Set NewSearchCommand.Config = Config
    Set NewSearchCommand.Tokens = Tokens
End Function

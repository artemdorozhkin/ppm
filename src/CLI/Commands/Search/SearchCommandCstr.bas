Attribute VB_Name = "SearchCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Search"
Option Explicit

Public Function NewSearchCommand(ByRef Tokens As Tokens) As SearchCommand
    Set NewSearchCommand = New SearchCommand
    Set NewSearchCommand.Tokens = Tokens
End Function

Attribute VB_Name = "HelpCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Help"
Option Explicit

Public Function NewHelpCommand(Optional ByRef Config As Config, Optional ByRef Tokens As Tokens) As HelpCommand
    Set NewHelpCommand = New HelpCommand
    Set NewHelpCommand.Config = Config
    Set NewHelpCommand.Tokens = Tokens
End Function

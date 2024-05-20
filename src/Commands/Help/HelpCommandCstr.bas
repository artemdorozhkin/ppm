Attribute VB_Name = "HelpCommandCstr"
'@Folder "PearPMProject.src.Commands.Help"
Option Explicit

Public Function NewHelpCommand(Optional ByRef Tokens As Tokens) As HelpCommand
    Set NewHelpCommand = New HelpCommand
    Set NewHelpCommand.Tokens = Tokens
End Function

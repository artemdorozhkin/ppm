Attribute VB_Name = "RefCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Ref"
Option Explicit

Public Function NewRefCommand(ByRef Tokens As Tokens) As RefCommand
    Set NewRefCommand = New RefCommand
    Set NewRefCommand.Tokens = Tokens
End Function

Attribute VB_Name = "BuildCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Build"
Option Explicit

Public Function NewBuildCommand(ByRef Tokens As Tokens) As BuildCommand
    Set NewBuildCommand = New BuildCommand
    Set NewBuildCommand.Tokens = Tokens
End Function

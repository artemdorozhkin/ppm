Attribute VB_Name = "AuthCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Auth"
Option Explicit

Public Function NewAuthCommand(ByRef Tokens As Tokens) As AuthCommand
    Set NewAuthCommand = New AuthCommand
    Set NewAuthCommand.Tokens = Tokens
End Function

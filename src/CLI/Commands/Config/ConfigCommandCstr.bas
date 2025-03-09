Attribute VB_Name = "ConfigCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Config"
Option Explicit

Public Function NewConfigCommand(ByRef Tokens As Tokens) As ConfigCommand
    Set NewConfigCommand = New ConfigCommand
    Set NewConfigCommand.Tokens = Tokens
End Function

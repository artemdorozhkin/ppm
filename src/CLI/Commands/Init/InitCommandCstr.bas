Attribute VB_Name = "InitCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Init"
Option Explicit

Public Function NewInitCommand(ByRef Tokens As Tokens) As InitCommand
    Set NewInitCommand = New InitCommand
    Set NewInitCommand.Tokens = Tokens
End Function

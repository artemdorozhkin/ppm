Attribute VB_Name = "InitCommandCstr"
'@Folder("PearPMProject.src.Commands.Init")
Option Explicit

Public Function NewInitCommand(ByRef Tokens As Tokens) As InitCommand
    Set NewInitCommand = New InitCommand
    Set NewInitCommand.Tokens = Tokens
End Function

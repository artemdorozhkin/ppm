Attribute VB_Name = "InitCommandCstr"
'@Folder("PearPMProject.src.Commands.Init")
Option Explicit

Public Function NewInitCommand(ByRef Args As Variant) As InitCommand
    Set NewInitCommand = New InitCommand
    NewInitCommand.Args = Args
End Function

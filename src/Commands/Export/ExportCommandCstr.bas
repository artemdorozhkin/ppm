Attribute VB_Name = "ExportCommandCstr"
'@Folder("PearPMProject.src.Commands.Export")
Option Explicit

Public Function NewExportCommand(ByRef Tokens As Tokens) As ExportCommand
    Set NewExportCommand = New ExportCommand
    Set NewExportCommand.Tokens = Tokens
End Function

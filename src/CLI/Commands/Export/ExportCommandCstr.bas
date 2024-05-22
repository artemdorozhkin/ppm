Attribute VB_Name = "ExportCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Export"
Option Explicit

Public Function NewExportCommand(ByRef Config As Config, ByRef Tokens As Tokens) As ExportCommand
    Set NewExportCommand = New ExportCommand
    Set NewExportCommand.Config = Config
    Set NewExportCommand.Tokens = Tokens
End Function

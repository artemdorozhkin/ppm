Attribute VB_Name = "ExportCommandCstr"
'@Folder("PearPMProject.src.Commands.Export")
Option Explicit

Public Function NewExportCommand(ByRef Args As Variant) As ExportCommand
    Set NewExportCommand = New ExportCommand
    NewExportCommand.Args = Args
End Function

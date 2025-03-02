Attribute VB_Name = "ModuleCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Module"
Option Explicit

Public Function NewModuleCommand(ByRef Config As Config, ByRef Tokens As Tokens) As ModuleCommand
    Set NewModuleCommand = New ModuleCommand
    Set NewModuleCommand.Config = Config
    Set NewModuleCommand.Tokens = Tokens
End Function

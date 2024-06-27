Attribute VB_Name = "SyncCommandCstr"
'@Folder "PearPMProject.src.CLI.Commands.Sync"
Option Explicit

Public Function NewSyncCommand(ByRef Config As Config, ByRef Tokens As Tokens) As SyncCommand
    Set NewSyncCommand = New SyncCommand
    Set NewSyncCommand.Config = Config
    Set NewSyncCommand.Tokens = Tokens
End Function

Attribute VB_Name = "SyncCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Sync"
Option Explicit

Public Function NewSyncCommand(ByRef Tokens As Tokens) As SyncCommand
    Set NewSyncCommand = New SyncCommand
    Set NewSyncCommand.Tokens = Tokens
End Function

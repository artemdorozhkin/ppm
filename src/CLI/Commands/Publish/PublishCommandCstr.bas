Attribute VB_Name = "PublishCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Publish"
Option Explicit

Public Function NewPublishCommand(ByRef Tokens As Tokens) As PublishCommand
    Set NewPublishCommand = New PublishCommand
    Set NewPublishCommand.Tokens = Tokens
End Function

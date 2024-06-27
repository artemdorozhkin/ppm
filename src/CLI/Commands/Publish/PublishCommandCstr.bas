Attribute VB_Name = "PublishCommandCstr"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Publish"
Option Explicit

Public Function NewPublishCommand(ByRef Config As Config, ByRef Tokens As Tokens) As PublishCommand
    Set NewPublishCommand = New PublishCommand
    Set NewPublishCommand.Config = Config
    Set NewPublishCommand.Tokens = Tokens
End Function

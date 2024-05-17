Attribute VB_Name = "HelpCommandCstr"
'@Folder "PearPMProject.src.Commands.Help"
Option Explicit

Public Function NewHelpCommand(ByRef Args As Variant) As HelpCommand
    Set NewHelpCommand = New HelpCommand
    NewHelpCommand.Args = Args
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstallCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Install"
Option Explicit

Implements ICommand

Private Type TInstallCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TInstallCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "install"
    this.CommandInfo.Description = "Installs the selected package."
    this.CommandInfo.Usage = Array("[<some package>@<some version>] [<options>]")
End Sub

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Immediate.WriteLine "ppm install command"
End Sub
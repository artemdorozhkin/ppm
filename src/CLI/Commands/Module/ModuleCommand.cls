VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Module"
Option Explicit

Implements ICommand

Private Type TConfigCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens

    Name As String
    Path As String
    FullPath As String
    Parts As Variant
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/module"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "module"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
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
    If Not this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        ppm "module --help"
        Exit Sub
    End If

    this.FullPath = Me.GetFullPath(this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text)
    this.Parts = Strings.Split(this.FullPath, ".")
    this.Name = this.Parts(UBound(this.Parts))

    If UBound(this.Parts) > 0 Then
        ReDim Preserve this.Parts(UBound(this.Parts) - 1)
        If this.Parts(0) = "" Then
            this.Parts(0) = SelectedProject.Name
        End If
        this.Path = Strings.Join(this.Parts, ".")
        this.Path = Strings.Replace(this.Path, "..", ".")
    Else
        this.Parts = Empty
    End If

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.GetFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If
    If IsFalse(SubCommandName) Or SubCommandName = "add" Then
        Me.AddModule
        Exit Sub
    End If

    If IsTrue(SubCommandName) Then
      #If DEV Then
        Dim Component As VBComponent
      #Else
        Dim Component As Object
      #End If
        Set Component = SelectedProject.GetComponent(Name:=this.Name)

        If IsFalse(Component) Then
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/moduleNotExists"), this.FullPath)
            Exit Sub
        End If

        Select Case SubCommandName
            Case "delete"
                SelectedProject.Components.Remove VBComponent:=Component
                Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/moduleDeleted"), this.FullPath)
                Exit Sub
            Case "move"
                Dim Module As Module
                Set Module = NewModule(Component:=Component)
                Module.Move ToPath:=this.Path
                Module.Name = this.Name
                Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/moduleMoved"), this.FullPath)
                Exit Sub
            Case Else
                ppm "module -- help"
        End Select
    End If
End Sub

Public Function GetFullPath(ByVal RawPath As String) As String
    Dim FullPath As String
    FullPath = Strings.Replace(RawPath, "/", ".")
    FullPath = Strings.Replace(FullPath, "\", ".")

    GetFullPath = FullPath
End Function

Public Sub AddModule()
    If SelectedProject.IsComponentExists(this.Name) Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/moduleAlreadyExists"), this.Name)
        Exit Sub
    End If

  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    Set Component = SelectedProject.AddComponent(Name:=this.Name)

    Dim Module As Module
    If Not Information.IsEmpty(this.Parts) Then
        Set Module = NewModule(Component:=Component)
        Module.Move ToPath:=this.Path
    End If

    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/moduleAdded"), this.FullPath)
End Sub

Attribute VB_Name = "CLI"
'@Folder "PearPMProject.src.CLI"
Option Explicit

Private Type TCLI
    Lang As Lang
End Type

Private this As TCLI

Public Property Get Lang() As Lang
    Set Lang = this.Lang
End Property

Public Property Get Commands() As Variant
    Commands = Array( _
        "config", _
        "export", _
        "help", _
        "init", _
        "install", _
        "publish", _
        "sync", _
        "uninstall", _
        "version" _
    )
End Property

Public Property Get SubCommands() As Variant
    SubCommands = Array( _
        "get", _
        "set", _
        "delete", _
        "list", _
        "edit", _
        "major", _
        "minor", _
        "patch" _
    )
End Property

#If DEV Then
  Public Property Get Aliases() As Dictionary
#Else
  Public Property Get Aliases() As Object
#End If
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If

    Buffer("i") = "install"
    Buffer("add") = "install"

    Buffer("exp") = "export"
    Buffer("save") = "export"

    Buffer("create") = "init"
    Buffer("new") = "init"

    Buffer("c") = "config"
    Buffer("cfg") = "config"

    Buffer("ls") = "list"

    Buffer("load") = "sync"

    Buffer("rm") = "uninstall"

    Set Aliases = Buffer
End Property

Public Function ParseCommand(ByRef Tokens As Tokens) As ICommand
    Dim Config As Config: Set Config = NewConfig(Definitions.Items, Tokens)
    If Tokens.Count = 0 Then
        Set ParseCommand = NewHelpCommand()
        Exit Function
    ElseIf Config.GetValue("help") Then
        Set ParseCommand = NewHelpCommand(Config, Tokens)
        Exit Function
    ElseIf Tokens.Count = 1 And Tokens.IncludeDefinition(Definitions("version")) Then
        ShowVersion
        End
    End If

    Dim CommandToken As SyntaxToken: Set CommandToken = Tokens(1) ' collection starts from 1
    If CommandToken.Kind <> TokenKind.Command Then
        Immediate.WriteLine "Unknown command " & CommandToken.Text
        End
    End If

    Dim CommandName As String: CommandName = CLI.FindCommand(CommandToken.Text)
    Dim Command As Variant: Set Command = Application.Run( _
        PStrings.FString("New{0}Command", CommandName), Config, Tokens _
    )
    Set ParseCommand = Command
End Function

Public Function FindCommand(ByVal Name As String) As String
    If PArray.IncludesAny(Commands, Name) Then
        FindCommand = Name
        Exit Function
    End If

    If Aliases.Exists(Name) Then
        FindCommand = Aliases(Name)
        Exit Function
    End If
End Function

Public Sub InitLang()
    Dim Config As Config: Set Config = NewConfig(Definitions.Items)
    Dim SelectedLang As String
    SelectedLang = Strings.LCase(Config.GetValue("language"))

    Set this.Lang = NewLang(SelectedLang)
End Sub

Private Sub ShowVersion()
    Dim Pack As Pack: Set Pack = NewPack(ThisProject.GetComponent("package"))
    Immediate.WriteLine Pack.Version
End Sub

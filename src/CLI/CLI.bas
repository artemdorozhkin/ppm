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
        "auth", _
        "class", _
        "config", _
        "export", _
        "help", _
        "init", _
        "install", _
        "module", _
        "publish", _
        "ref", _
        "search", _
        "sync", _
        "uninstall", _
        "version" _
    )
End Property

Public Property Get SubCommands() As Variant
    SubCommands = Array( _
        "delete", _
        "edit", _
        "generate", _
        "get", _
        "list", _
        "major", _
        "minor", _
        "move", _
        "patch", _
        "set", _
        "update" _
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

    Buffer("cls") = "class"

    Buffer("c") = "config"
    Buffer("cfg") = "config"

    Buffer("exp") = "export"
    Buffer("save") = "export"

    Buffer("create") = "init"
    Buffer("new") = "init"

    Buffer("i") = "install"
    Buffer("add") = "install"

    Buffer("ls") = "list"

    Buffer("m") = "module"
    Buffer("bas") = "module"
    Buffer("mv") = "move"

    Buffer("r") = "ref"
    Buffer("references") = "ref"

    Buffer("find") = "search"
    Buffer("s") = "search"

    Buffer("load") = "sync"

    Buffer("rm") = "uninstall"

    Buffer("upd") = "update"

    Set Aliases = Buffer
End Property

Public Function ParseCommand(ByRef Tokens As Tokens) As ICommand
    Dim Config As Config: Set Config = NewConfig(ConfigScopes.DefaultScope)
    If Tokens.Count = 0 Then
        Set ParseCommand = NewHelpCommand()
        Exit Function
    ElseIf Tokens.IncludeDefinition(Definitions("help")) Then
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
    Dim Config As Config: Set Config = NewConfig(ConfigScopes.UserScope)
    Dim SelectedLang As String
    SelectedLang = Strings.LCase(GetFirstTrue( _
        Config.GetValue("language"), _
        Definitions("language").Default _
    ))

    Set this.Lang = NewLang(SelectedLang)
End Sub

Private Sub ShowVersion()
    Dim Pack As Pack: Set Pack = NewPack(ThisProject.GetComponent("package"))
    Immediate.WriteLine Pack.Version
End Sub

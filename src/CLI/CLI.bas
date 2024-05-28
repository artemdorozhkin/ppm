Attribute VB_Name = "CLI"
'@Folder "PearPMProject.src.CLI"
Option Explicit

Public Property Get Commands() As Variant
    Commands = Array( _
        "config", _
        "export", _
        "help", _
        "init", _
        "install", _
        "publish" _
    )
End Property

Public Property Get SubCommands() As Variant
    SubCommands = Array( _
        "get", _
        "set", _
        "delete", _
        "list", _
        "edit" _
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
    Buffer("rm") = "delete"

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

    Dim Command As String: Command = CLI.FindCommand(CommandToken.Text)
    Set ParseCommand = Application.Run( _
        PStrings.FString("New{0}Command", Command), Config, Tokens _
    )
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

Private Sub ShowVersion()
    Dim Project As Project
    Dim VBProject As VBProject
    For Each VBProject In Application.VBE.VBProjects
        If PStrings.IsEqual(VBProject.Name, "PearPMProject") Then
            Set Project = NewProject(VBProject)
            Exit For
        End If
    Next

    Dim Pack As Pack: Set Pack = NewPack(Project)
    Pack.Read

    Immediate.WriteLine Pack.Version
End Sub

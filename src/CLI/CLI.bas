Attribute VB_Name = "CLI"
'@Folder "PearPMProject.src.CLI"
Option Explicit

Public Property Get Commands() As Variant
    Commands = Array( _
        "config", _
        "export", _
        "help", _
        "init", _
        "install" _
    )
End Property

Public Property Get Definitions()
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary(vbTextCompare)
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary(vbTextCompare)
  #End If
    Dim Definition As Definition

    Set Definition = NewDefinition( _
        Key:="_after-dialog", _
        KeyType:=vbBoolean, _
        Short:="", _
        Default:=True _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="encoding", _
        KeyType:=vbString, _
        Short:="e", _
        Default:="UTF-8", _
        Description:="\\t\\tExport files with set encoding." _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="help", _
        KeyType:=vbBoolean, _
        Short:="h", _
        Default:=True, _
        Description:="\\t\\tShow help about command." _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="name", _
        KeyType:=vbString, _
        Short:="n", _
        Default:="", _
        Description:="\\tSets name." _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="save-struct", _
        KeyType:=vbBoolean, _
        Short:="s", _
        Default:=False, _
        Description:="\\tSave the RubberDuck structure when exporting a project." _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definition = NewDefinition( _
        Key:="yes", _
        KeyType:=vbBoolean, _
        Short:="y", _
        Default:=True, _
        Description:="\\tSkips dialog and sets default values." _
    )
    Set Buffer(Definition.Key) = Definition

    Set Definitions = Buffer
End Property

#If DEV Then
  Public Property Get Aliases() As Dictionary
#Else
  Public Property Get Aliases() As Object
#End If
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary(vbTextCompare)
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary(vbTextCompare)
  #End If

    Buffer("i") = "install"
    Buffer("add") = "install"

    Buffer("exp") = "export"
    Buffer("save") = "export"

    Buffer("create") = "init"
    Buffer("new") = "init"

    Buffer("c") = "config"
    Buffer("cfg") = "config"

    Set Aliases = Buffer
End Property

Public Function ParseCommand(ByRef Tokens As Tokens) As ICommand
    If Tokens.Count = 0 Then
        Set ParseCommand = NewHelpCommand()
        Exit Function
    ElseIf Tokens.IncludeToken("h", ShortOptionItem) Or _
           Tokens.IncludeToken("help", OptionItem) Then
        Set ParseCommand = NewHelpCommand(Tokens)
        Exit Function
    End If

    Dim CommandToken As SyntaxToken: Set CommandToken = Tokens(1) ' collection starts from 0
    If CommandToken.Kind <> TokenKind.Command Then
        Immediate.WriteLine "Unknown command " & CommandToken.Text
        End
    End If

    Dim Command As String: Command = CLI.FindCommand(CommandToken.Text)
    Set ParseCommand = Application.Run(FString("New{0}Command", Command), Tokens)
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands"
Option Explicit

Public Name As String
Public Description As String
#If DEV Then
  Public Params As Dictionary
#Else
  Public Params As Object
#End If
Public Usage As Variant

Private Sub Class_Initialize()
    Set Me.Params = NewDictionary(VbCompareMethod.vbTextCompare)
End Sub

Public Function GetUsage() As String
    Dim Usage As String
    Usage = PStrings.FString(Me.Description & "\\n")
    Usage = Usage & PStrings.FString("\\nUsage:\\nppm ""{0} {1}""\\n", Me.Name, Strings.Join(Me.Usage, "\\n  "))

    If Me.Params.Count > 0 Then
        Usage = Usage & "\\nOptions:\\n"
        Dim Param As Variant
        For Each Param In Me.Params
            Dim Definition As Definition: Set Definition = Definitions(Param)
            Dim Value As String: Value = Empty
            If Strings.Len(Definition.Short) > 0 Then
                Value = PStrings.FString("-{0}|", Definition.Short)
            End If
            Value = Value & FString("--{0}", Definition.Key)
            Usage = Usage & PStrings.FString( _
                "{0}{1}\\n", Value, Params(Param) _
            )
        Next
    End If

    Dim CommandAliases As Collection: Set CommandAliases = New Collection
    Dim CLIAliases As Object: Set CLIAliases = CLI.Aliases
    Dim Alias As Variant
    For Each Alias In CLIAliases
        If IsEqual(CLIAliases(Alias), Me.Name) Then
            CommandAliases.Add Alias
        End If
    Next

    If CommandAliases.Count > 0 Then
        Usage = Usage & PStrings.FString("\\nAliases:\\n{0}\\n", Strings.Join(PCollection.ToArray(CommandAliases), ", "))
    End If

    GetUsage = PStrings.FString(Usage)
End Function

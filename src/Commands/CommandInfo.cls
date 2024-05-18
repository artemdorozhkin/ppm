VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Commands"
Option Explicit

Public Name As String
Public Description As String
Public Params As Variant
Public Usage As Variant

Private Sub Class_Initialize()
    Set Me.Params = NewDictionary()
End Sub

Public Function GetUsage() As String
    Dim Usage As String
    Usage = FString(Me.Description & "\\n")
    Usage = Usage & FString("\\nUsage:\\nppm ""{0} {1}""\\n", Me.Name, Strings.Join(Me.Usage, "\\n  "))

    If Me.Params.Count > 0 Then
        Usage = Usage & "\\nOptions:\\n"
        Dim Param As Variant
        For Each Param In Me.Params
            Dim About As String: About = Me.Params(Param)
            Usage = Usage & FString("--{0}{1}\\n", Param, About)
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
        Usage = Usage & FString("\\nAliases:\\n{0}\\n", Strings.Join(ToArray(CommandAliases), ", "))
    End If

    GetUsage = FString(Usage)
End Function

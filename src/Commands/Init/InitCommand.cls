VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Commands.Init"
Option Explicit

Implements ICommand

Private Type TInitCommand
    Args As Variant
    CommandInfo As CommandInfo
End Type

Private this As TInitCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "init"
    this.CommandInfo.Description = "Create a package module."
    this.CommandInfo.Params.Add "yes", "\\tSkip dialog and sets package by default."
    this.CommandInfo.Usage = Array( _
        "<package name>" _
    )
End Sub

Public Property Let Args(ByVal RHS As Variant)
    this.Args = RHS
End Property

Public Function FindProjectName() As String
    this.Args = Strings.Filter(this.Args, "init", include:=False, Compare:=vbTextCompare)
    this.Args = Strings.Filter(this.Args, "--yes", include:=False, Compare:=vbTextCompare)
    this.Args = Strings.Filter(this.Args, "-y", include:=False, Compare:=vbTextCompare)
    If UBound(this.Args) > 0 Then Err.Raise 9, TypeName(Me), "Invalid count of options."
    If UBound(this.Args) = -1 Then Exit Function
    FindProjectName = this.Args(0)
End Function

Public Sub CreatePack()
    If SelectedProject.IsModuleExists("package") Then
        Immediate.WriteLine "module already exists"
        Exit Sub
    End If

    Dim PackModule As VBComponent: Set PackModule = SelectedProject.AddModule("package")
    With PackModule.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString FString( _
            "' {\\n" & _
            "'   ""name"": ""{0}"",\\n" & _
            "'   ""version"": ""{1}"",\\n" & _
            "'   ""description"": ""{2}"",\\n" & _
            "'   ""author"": ""{3}"",\\n" & _
            "'   ""git"": ""{4}""\\n" & _
            "' }", _
            PackInfo.Name, PackInfo.Version, PackInfo.Description, PackInfo.Author, PackInfo.GitURL _
        )
    End With
End Sub

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim SkipByDefault As Boolean
    SkipByDefault = PArray.IncludesAny(this.Args, "-y", "--yes")

    If (UBound(this.Args) > 0 And Not SkipByDefault) Or _
       (UBound(this.Args) > 1 And SkipByDefault) Then
        InitDialog.SetNameAndDefault Me.FindProjectName()
    ElseIf Not SkipByDefault Then
        InitDialog.SetDefault
    End If

    If SkipByDefault Then
        Dim ms As Double: ms = DateTime.Timer
        Me.CreatePack
        ms = DateTime.Timer - ms
        Immediate.WriteLine FString("package is initialized in {0}", ConvertTime(ms))
    Else
        InitDialog.Start
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuildCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Build"
Option Explicit

Implements ICommand

Private Type TBuildCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TBuildCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/build"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "build"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("output") = CLI.Lang.GetValue("param", "name=output")
    this.CommandInfo.Params.Item("file-format") = CLI.Lang.GetValue("param", "name=file-format")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If Not SelectedProject.IsSaved Then
        Immediate.WriteLine "ERR: Project must be saved first"
        Exit Sub
    End If

    Dim CurrentPath As String
    CurrentPath = SelectedProject.Path

    Dim CWD As String
    CWD = FileSystem.CurDir()

    FileSystem.ChDir SelectedProject.Folder

    Dim AllowedExtensions() As String
    AllowedExtensions = Strings.Split("xlsb,xlsm,xls,xla,xlam", ",")

    Dim FilePath As String
    Dim Output As Definition
    Set Output = Definitions("output")

    Dim FileFormat As Definition
    Set FileFormat = Definitions("file-format")
    Dim Extension As String
    Extension = "xlsb"
    If this.Tokens.IncludeDefinition(FileFormat) Then
        Extension = this.Tokens.PopTokenByDefinition(FileFormat).Text
    End If

    If Not PArray.IncludesAny(AllowedExtensions, Extension) Then
        Immediate.WriteLine "ERR: Not supported extension:", PFileSystem.GetFileExt(FilePath)
        Exit Sub
    End If

    If this.Tokens.IncludeDefinition(Output) Then
        FilePath = this.Tokens.PopTokenByDefinition(Output).Text
        If Strings.Len(PFileSystem.GetFileExt(FilePath)) = 0 Then
            FilePath = PStrings.FString("{0}.{1}", FilePath, Extension)
        End If
    Else
        FilePath = Me.BuildFilePath(Extension)
    End If

    Me.RemoveDevsAttributes
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Workbooks(FileSystem.Dir(SelectedProject.Path)).SaveAs FilePath
    Application.Workbooks(FileSystem.Dir(FilePath)).Close SaveChanges:=False
    Application.Workbooks.Open CurrentPath
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    FileSystem.ChDir CWD
    Immediate.WriteLine "Successfully builded: ", FilePath
End Sub

Public Function BuildFilePath(ByVal Extension As String) As String
    Dim Name As String
    Dim Version As String
    If SelectedProject.IsComponentExists("package") Then
        Dim Pack As Pack
        Set Pack = NewPack(SelectedProject.GetComponent("package"))
        Name = Pack.Name
        Version = Pack.Version
    Else
        Name = SelectedProject.Name
        Version = "build"
    End If

    If IsFalse(Name) And IsFalse(Version) Then
        BuildFilePath = PStrings.FString( _
            "build{0}.{1}", _
            Strings.Format(DateTime.Now, "ddmmyy_hhnnss"), _
            Extension _
        )
    Else
        BuildFilePath = PStrings.FString( _
            "{0}@{1}.{2}", _
            Name, _
            Version, _
            Extension _
        )
    End If
End Function

Public Sub RemoveDevsAttributes()
    If Not SelectedProject.IsComponentExists("package") Then Exit Sub

    Dim Pack As Pack
    Set Pack = NewPack(SelectedProject.GetComponent("package"))
    Dim JSON As Object
    Set JSON = Pack.JSON
    If IsTrue(Pack.DevDependencies) Then
        If Pack.DevDependencies.Count > 0 Then
            ppm "uninstall " & Strings.Join(Pack.DevDependencies.Keys(), " ")
        End If
    End If

    If IsTrue(Pack.DevReferences) Then
        If Pack.DevReferences.Count > 0 Then
            ppm "ref delete " & Strings.Join(Pack.DevReferences.Keys(), " ")
        End If
    End If

    Set Pack.JSON = JSON
    Pack.Update
End Sub

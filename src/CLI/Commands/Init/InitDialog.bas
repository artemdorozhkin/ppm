Attribute VB_Name = "InitDialog"
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.CLI.Commands.Init"
Option Explicit

Public Type TPackageInfo
    Name As String
    Version As String
    Author As String
    Description As String
    GitURL As String
End Type

Public PackInfo As TPackageInfo

Public Sub Start()
    Dim Default As String: Default = GetOrDefault(PackInfo.Name, "name")
    Immediate.ReadLine PStrings.FString("project name ({0}):", Default), "InitDialog.Step2"
End Sub

Public Sub Step2(Optional ByVal ProjectName As String)
    If IsFalse(PackInfo.Name) And IsFalse(ProjectName) Then
        PackInfo.Name = GetOrDefault(ProjectName, "name")
    ElseIf IsTrue(ProjectName) Then
        PackInfo.Name = ProjectName
    End If

    Dim Default As String: Default = GetOrDefault("", "version")
    Immediate.ReadLine PStrings.FString("version ({0}):", Default), "InitDialog.Step3"
End Sub

Public Sub Step3(Optional ByVal Version As String)
    PackInfo.Version = GetOrDefault(Version, "version")
    Immediate.ReadLine "description:", "InitDialog.Step4"
End Sub

Public Sub Step4(Optional ByVal Description As String)
    PackInfo.Description = Description
    Dim Default As String: Default = GetOrDefault("", "author-name")
    Immediate.ReadLine PStrings.FString("author ({0}):", Default), "InitDialog.Step5"
End Sub

Public Sub Step5(Optional ByVal Author As String)
    PackInfo.Author = GetOrDefault(Author, "author-name")
    Dim Default As String: Default = GetOrDefault("", "author-url")
    Immediate.ReadLine PStrings.FString("git repository ({0}):", Default), "InitDialog.Step6"
End Sub

Public Sub Step6(Optional ByVal GitURL As String)
    PackInfo.GitURL = GetOrDefault(GitURL, "author-url")

    Immediate.WriteLine
    With PackInfo
        Dim Info As String
        Info = PStrings.FString( _
            "name: {0}\\nversion: {1}\\ndescription: {2}\\nauthor: {3}\\ngit: {4}\\n", _
            .Name, .Version, .Description, .Author, .GitURL _
        )
    End With

    Immediate.WriteLine Info
    Immediate.ReadLine "is this ok? (Y/n):", "InitDialog.Step7"
End Sub

Public Sub Step7(Optional ByVal Confirm As String = "y")
    If IsEqual(Confirm, "y") Then
        ppm "init -y --_after-dialog"
    Else
        Immediate.WriteLine "Aborted"
    End If
End Sub

Public Function GetOrDefault(ByVal Value As String, ByVal ConfigKey As String) As String
    If Strings.Len(Value) > 0 Then
        GetOrDefault = Value
        Exit Function
    End If

    Dim Config As Config
    Set Config = NewConfig(Definitions.Items)
    GetOrDefault = Config.GetValue(ConfigKey)
End Function

Public Function SetNameAndDefault(ByVal Name As String) As TPackageInfo
    PackInfo = SetDefault()
    PackInfo.Name = GetOrDefault(Name, "name")
    SetNameAndDefault = PackInfo
End Function

Public Function SetDefault() As TPackageInfo
    With PackInfo
        .Name = GetOrDefault("", "name")
        .Version = GetOrDefault("", "version")
        .Description = ""
        .Author = GetOrDefault("", "author-name")
        .GitURL = GetOrDefault("", "author-url")
    End With

    SetDefault = PackInfo
End Function

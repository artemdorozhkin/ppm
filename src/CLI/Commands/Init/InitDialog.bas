Attribute VB_Name = "InitDialog"
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
    Dim Default As String: Default = GetOrDefault(PackInfo.Name, "ProjectName")
    Immediate.ReadLine FString("project name ({0}):", Default), "InitDialog.Step2"
End Sub

Public Sub Step2(Optional ByVal ProjectName As String)
    Dim Name As String
    If Strings.Len(PackInfo.Name) = 0 Then
        PackInfo.Name = GetOrDefault(ProjectName, "ProjectName")
    End If

    Dim Default As String: Default = GetOrDefault("", "Version")
    Immediate.ReadLine FString("version ({0}):", Default), "InitDialog.Step3"
End Sub

Public Sub Step3(Optional ByVal Version As String)
    PackInfo.Version = GetOrDefault(Version, "Version")
    Immediate.ReadLine "description:", "InitDialog.Step4"
End Sub

Public Sub Step4(Optional ByVal Description As String)
    PackInfo.Description = GetOrDefault(Description, "Description")

    Dim Default As String: Default = GetOrDefault("", "ProjectName")
    Immediate.ReadLine "author:", "InitDialog.Step5"
End Sub

Public Sub Step5(Optional ByVal Author As String)
    PackInfo.Author = GetOrDefault(Author, "Author")
    Immediate.ReadLine "git repository:", "InitDialog.Step6"
End Sub

Public Sub Step6(Optional ByVal GitURL As String)
    PackInfo.GitURL = GetOrDefault(GitURL, "GitURL")

    Immediate.WriteLine
    With PackInfo
        Dim Info As String
        Info = FString( _
            "name: {0}\\nversion: {1}\\ndescription: {2}\\nauthor: {3}\\ngit: {4}\\n", _
            .Name, .Version, .Description, .Author, .GitURL _
        )
    End With

    Immediate.WriteLine Info
    Immediate.ReadLine "is this ok? (Y/n):", "InitDialog.Step7"
End Sub

Public Sub Step7(Optional ByVal Confirm As String)
    If IsEqual(Confirm, "n") Then
        Immediate.WriteLine "Aborted"
    Else
        ppm "init -y --_after-dialog"
    End If
End Sub

Public Function GetOrDefault(ByVal Value As String, ByVal ConfigKey As String) As String
    If Strings.Len(Value) > 0 Then
        GetOrDefault = Value
        Exit Function
    End If

    Static Section As Dictionary
    If IsFalse(Section) Then
        Dim Config As ConfigIO
        Set Config = NewConfigIO(CLI.GetGlobalConfigPath())
        Set Section = Config.ReadSection("Init")
    End If
    GetOrDefault = Section(ConfigKey)
End Function

Public Function SetNameAndDefault(ByVal Name As String) As TPackageInfo
    PackInfo = SetDefault()
    PackInfo.Name = GetOrDefault(Name, "ProjectName")
    SetNameAndDefault = PackInfo
End Function

Public Function SetDefault() As TPackageInfo
    With PackInfo
        .Name = GetOrDefault("", "ProjectName")
        .Version = GetOrDefault("", "Version")
        .Description = ""
        .Author = ""
        .GitURL = ""
    End With

    SetDefault = PackInfo
End Function


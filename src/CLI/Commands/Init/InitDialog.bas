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

Public InitPack As TPackageInfo
Public PackInfo As TPackageInfo

Public Sub Start()
    Immediate.ReadLine PStrings.FString("project name ({0}):", InitPack.Name), "InitDialog.Step2"
End Sub

Public Sub Step2(Optional ByVal ProjectName As String)
    PackInfo.Name = GetFirstTrue(ProjectName, InitPack.Name)
    Immediate.ReadLine PStrings.FString("version ({0}):", InitPack.Version), "InitDialog.Step3"
End Sub

Public Sub Step3(Optional ByVal Version As String)
    PackInfo.Version = GetFirstTrue(Version, InitPack.Version)
    Immediate.ReadLine PStrings.FString("description ({0}):", InitPack.Description), "InitDialog.Step4"
End Sub

Public Sub Step4(Optional ByVal Description As String)
    PackInfo.Description = GetFirstTrue(Description, InitPack.Description)
    Immediate.ReadLine PStrings.FString("author ({0}):", InitPack.Author), "InitDialog.Step5"
End Sub

Public Sub Step5(Optional ByVal Author As String)
    PackInfo.Author = GetFirstTrue(Author, InitPack.Author)
    Immediate.ReadLine PStrings.FString("git repository ({0}):", InitPack.Author), "InitDialog.Step6"
End Sub

Public Sub Step6(Optional ByVal GitURL As String)
    PackInfo.GitURL = GetFirstTrue(GitURL, InitPack.GitURL)

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

Private Function GetFirstTrue(ByVal Value1 As String, ByVal Value2 As String) As String
    GetFirstTrue = Interaction.IIf(IsTrue(Value1), Value1, Value2)
End Function

Public Function SetNameAndDefault(ByVal Name As String) As TPackageInfo
    PackInfo = SetDefault()
    PackInfo.Name = GetOrDefault(Name, "name")
    SetNameAndDefault = PackInfo
End Function

Public Function SetDefault( _
    Optional ByVal Name As Variant, _
    Optional ByVal Version As Variant, _
    Optional ByVal Description As Variant, _
    Optional ByVal Author As Variant, _
    Optional ByVal GitURL As Variant _
) As TPackageInfo
    With InitPack
        If Information.IsMissing(Name) Then
            .Name = GetOrDefault("", "name")
        Else
            .Name = Name
        End If

        If Information.IsMissing(Version) Then
            .Version = GetOrDefault("", "version")
        Else
            .Version = Version
        End If

        If Information.IsMissing(Description) Then
            .Description = ""
        Else
            .Description = Description
        End If

        If Information.IsMissing(Author) Then
            .Author = GetOrDefault("", "author-name")
        Else
            .Author = Author
        End If

        If Information.IsMissing(GitURL) Then
            .GitURL = GetOrDefault("", "author-url")
        Else
            .GitURL = GitURL
        End If
    End With

    SetDefault = PackInfo
End Function

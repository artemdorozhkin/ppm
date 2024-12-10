VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UninstallCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Uninstall"
Option Explicit

Implements ICommand

Private Type TUninstallCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
    Pack As Pack
End Type

Private this As TUninstallCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/uninstall"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "uninstall"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim ms As Double: ms = DateTime.Timer
    Immediate.WriteLine CLI.Lang.GetValue("messages/uninstalling")

    Set this.Pack = NewPack(SelectedProject.GetComponent("package"))

  #If DEV Then
    Dim Libs As Dictionary: Set Libs = GetAllInstalledLibs()
    Dim RmDepenpencies As Dictionary: Set RmDepenpencies = GetRmDependencies()
  #Else
    Dim Libs As Object: Set Libs = GetAllInstalledLibs()
    Dim RmDepenpencies As Object: Set RmDepenpencies = GetRmDependencies()
  #End If
    Dim NotUsed As Collection: Set NotUsed = New Collection

    If IsTrue(Libs) And IsTrue(RmDepenpencies) Then
        Dim Lib As Variant
        For Each Lib In Libs
            If Libs(Lib) > 1 Then GoTo Continue
            If Not RmDepenpencies.Exists(Lib) Then GoTo Continue
            NotUsed.Add Lib
Continue:
        Next
    End If

    If NotUsed.Count = 0 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/cantRemoveUsingDependencies")
        End
    End If

    Dim Dep As Variant
    For Each Dep In NotUsed
        RemoveDependency Dep
        this.Pack.Dependencies.Remove Dep
    Next
    this.Pack.Update

    ms = DateTime.Timer - ms
    Immediate.WriteLine PStrings.FString( _
        CLI.Lang.GetValue("messages/uninstalled"), _
        NotUsed.Count, _
        Utils.ConvertTime(ms) _
    )
End Sub

#If DEV Then
  Public Function GetAllInstalledLibs() As Dictionary
#Else
  Public Function GetAllInstalledLibs() As Object
#End If
  #If DEV Then
    Dim Libs As Dictionary: Set Libs = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim Libs As Object: Set Libs = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If
    CollectAndCountDependencies "package", Libs
    Set GetAllInstalledLibs = Libs
End Function

#If DEV Then
  Public Sub CollectAndCountDependencies(ByVal PackName As String, ByRef Output As Dictionary)
#Else
  Public Sub CollectAndCountDependencies(ByVal PackName As String, ByRef Output As Object)
#End If
    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent(PackName))
    If IsFalse(Pack.Dependencies) Then Exit Sub

    Dim Name As Variant
    For Each Name In Pack.Dependencies
        Output(Name) = Output(Name) + 1
        CollectAndCountDependencies Name & "Package", Output
    Next
End Sub

#If DEV Then
  Public Function GetRmDependencies() As Dictionary
#Else
  Public Function GetRmDependencies() As Object
#End If
    If IsFalse(this.Pack.Dependencies) Then Exit Function

  #If DEV Then
    Dim RmLibs As Dictionary: Set RmLibs = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim RmLibs As Object: Set RmLibs = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If

    Dim LibName As Variant
    For Each LibName In this.Tokens.GetTokensKind(TokenKind.Identifier)
        If this.Pack.Dependencies.Exists(LibName.Text) Then
            RmLibs(LibName.Text) = RmLibs(LibName.Text) + 1
            CollectAndCountDependencies LibName.Text & "Package", RmLibs
        End If
    Next

    Set GetRmDependencies = RmLibs
End Function

Public Sub RemoveDependency(ByVal Name As String)
    Dim Pack As Pack: Set Pack = NewPack(SelectedProject.GetComponent(Name & "Package"))

    Dim RootPath As String: RootPath = Pack.Root
  #If DEV Then
    Dim Components As VBComponents: Set Components = SelectedProject.Components
    Dim Component As VBComponent
  #Else
    Dim Components As Object: Set Components = SelectedProject.Components
    Dim Component As Object
  #End If

    For Each Component In Components
        Dim Module As Module: Set Module = NewModule(Component)
        If PStrings.StartsWith(Module.Path, RootPath, vbTextCompare) Then
            Components.Remove Component
        End If
    Next
End Sub

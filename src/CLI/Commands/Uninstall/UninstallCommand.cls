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
    Tokens As Tokens
    Pack As Pack
    HasDeps As Boolean
    HasDevDeps As Boolean
End Type

Private this As TUninstallCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/uninstall"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "uninstall"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

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
    this.HasDeps = IsTrue(this.Pack.Dependencies)
    this.HasDevDeps = IsTrue(this.Pack.DevDependencies)

  #If DEV Then
    Dim Libs As Dictionary: Set Libs = GetAllInstalledLibs()
    Dim RmDepenpencies As Dictionary: Set RmDepenpencies = GetRmDependencies()
  #Else
    Dim Libs As Object: Set Libs = GetAllInstalledLibs()
    Dim RmDepenpencies As Object: Set RmDepenpencies = GetRmDependencies()
  #End If
    If Libs.Count = 0 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/noInstalledLibs")
        Exit Sub
    End If

    If RmDepenpencies.Count = 0 And this.Tokens.Count = 1 Then
        Set RmDepenpencies = Libs
    ElseIf RmDepenpencies.Count = 0 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/nothingToRemove")
        Exit Sub
    End If

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

    Dim Has
    Dim Dep As Variant
    For Each Dep In NotUsed
        RemoveDependency Dep
        If this.HasDeps Then
            If this.Pack.Dependencies.Exists(Dep) Then
                this.Pack.Dependencies.Remove Dep
            End If
        End If

        If this.HasDevDeps Then
            If this.Pack.DevDependencies.Exists(Dep) Then
                this.Pack.DevDependencies.Remove Dep
            End If
        End If
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
    Dim HasDeps As Boolean
    HasDeps = IsTrue(Pack.Dependencies)
    Dim HasDevDeps As Boolean
    HasDevDeps = IsTrue(Pack.DevDependencies)
    If HasDeps Then
        Dim Name As Variant
        For Each Name In Pack.Dependencies
            Output(Name) = Output(Name) + 1
            CollectAndCountDependencies Name & "Package", Output
        Next
    End If
    
    If HasDevDeps Then
        For Each Name In Pack.DevDependencies
            Output(Name) = Output(Name) + 1
            CollectAndCountDependencies Name & "Package", Output
        Next
    End If
End Sub

#If DEV Then
  Public Function GetRmDependencies() As Dictionary
#Else
  Public Function GetRmDependencies() As Object
#End If
    If Not this.HasDeps And Not this.HasDevDeps Then
        Set GetRmDependencies = NewDictionary()
        Exit Function
    End If

  #If DEV Then
    Dim RmLibs As Dictionary: Set RmLibs = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim RmLibs As Object: Set RmLibs = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If

    Dim LibName As Variant
    For Each LibName In this.Tokens.GetTokensKind(TokenKind.Identifier)
        If this.HasDeps Then
            If this.Pack.Dependencies.Exists(LibName.Text) Then
                RmLibs(LibName.Text) = RmLibs(LibName.Text) + 1
                CollectAndCountDependencies LibName.Text & "Package", RmLibs
                GoTo Continue
            End If
        End If

        If this.HasDevDeps Then
            If this.Pack.DevDependencies.Exists(LibName.Text) Then
                RmLibs(LibName.Text) = RmLibs(LibName.Text) + 1
                CollectAndCountDependencies LibName.Text & "Package", RmLibs
                GoTo Continue
            End If
        End If

        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cantFindDep"), LibName.Text)
Continue:
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

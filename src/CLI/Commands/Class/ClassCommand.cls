VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Class"
Option Explicit

Implements ICommand

Private Type TConfigCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens

    Name As String
    CstrName As String
    Path As String
    FullPath As String
    Parts As Variant
End Type

Private this As TConfigCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/class"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "class"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("create-constructor") = CLI.Lang.GetValue("param", "name=create-constructor")
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
    If Not this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        ppm "class --help"
        Exit Sub
    End If

    this.FullPath = Me.GetFullPath(this.Tokens.GetFirstTokenKind(TokenKind.Identifier).Text)
    this.Parts = Strings.Split(this.FullPath, ".")
    this.Name = this.Parts(UBound(this.Parts))
    this.CstrName = this.Name & "Cstr"

    If UBound(this.Parts) > 0 Then
        ReDim Preserve this.Parts(UBound(this.Parts) - 1)
        If this.Parts(0) = "" Then
            this.Parts(0) = SelectedProject.Name
        End If
        this.Path = Strings.Join(this.Parts, ".")
        this.Path = Strings.Replace(this.Path, "..", ".")
    Else
        this.Parts = Empty
    End If

    Dim SubCommand As SyntaxToken
    Set SubCommand = this.Tokens.GetFirstTokenKind(TokenKind.SubCommand)
    Dim SubCommandName As String
    If IsTrue(SubCommand) Then
        SubCommandName = SubCommand.Text
    End If

    If CLI.Aliases.Exists(SubCommandName) Then
        SubCommandName = CLI.Aliases(SubCommandName)
    End If
    If IsFalse(SubCommandName) Or SubCommandName = "add" Then
        Me.AddClass
        If this.Tokens.IncludeDefinition(Definitions("create-constructor")) Then
            Me.AddConstructor
        End If
        Exit Sub
    End If

    If IsTrue(SubCommandName) Then
      #If DEV Then
        Dim Component As VBComponent
        Dim CstrComponent As VBComponent
      #Else
        Dim Component As Object
        Dim CstrComponent As Object
      #End If
        Set Component = SelectedProject.GetComponent(Name:=this.Name)

        If SelectedProject.IsComponentExists(this.CstrName) Then
            Set CstrComponent = SelectedProject.GetComponent(Name:=this.CstrName)
        End If

        If IsFalse(Component) Then
            Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/classNotExists"), this.FullPath)
            Exit Sub
        End If

        Select Case SubCommandName
            Case "delete"
                SelectedProject.Components.Remove VBComponent:=Component
                Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/classDeleted"), this.FullPath)
                If IsTrue(CstrComponent) Then
                    SelectedProject.Components.Remove VBComponent:=CstrComponent
                    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cstrDeleted"), this.FullPath)
                End If
                Exit Sub
            Case "move"
                Dim ClassModule As Module
                Set ClassModule = NewModule(Component:=Component)
                ClassModule.Move ToPath:=this.Path
                ClassModule.Name = this.Name
                Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/classMoved"), this.FullPath)
                If IsTrue(CstrComponent) Then
                    Dim CstrModule As Module
                    Set CstrModule = NewModule(Component:=CstrComponent)
                    CstrModule.Move ToPath:=this.Path
                    CstrModule.Name = this.CstrName
                    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cstrMoved"), this.FullPath)
                End If
                Exit Sub
            Case Else
                ppm "class -- help"
        End Select
    End If
End Sub

Public Function GetFullPath(ByVal RawPath As String) As String
    Dim FullPath As String
    FullPath = Strings.Replace(RawPath, "/", ".")
    FullPath = Strings.Replace(FullPath, "\", ".")

    GetFullPath = FullPath
End Function

Public Sub AddClass()
    If SelectedProject.IsComponentExists(this.Name) Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/classAlreadyExists"), this.Name)
        Exit Sub
    End If

  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    Set Component = SelectedProject.AddComponent(Name:=this.Name, Kind:=ModuleType.ClassModule)

    Dim ClassModule As Module
    If Not Information.IsEmpty(this.Parts) Then
        Set ClassModule = NewModule(Component:=Component)
        ClassModule.Move ToPath:=this.Path
    End If

    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/classAdded"), this.FullPath)
End Sub

Public Sub AddConstructor()
    If SelectedProject.IsComponentExists(this.CstrName) Then
        Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cstrAlreadyExists"), this.Name)
        Exit Sub
    End If

  #If DEV Then
    Dim Component As VBComponent
  #Else
    Dim Component As Object
  #End If
    Set Component = SelectedProject.AddComponent(Name:=this.CstrName)

    Dim CstrModule As Module
    If Not Information.IsEmpty(this.Parts) Then
        Set CstrModule = NewModule(Component:=Component)
        CstrModule.Move ToPath:=this.Path
        Dim Code As Variant: Code = Array( _
            "\\nPublic Function New{0}() As {0}", _
            "\\tSet New{0} = New {0}", _
            "End Function" _
        )
        CstrModule.Code = CstrModule.Code & PStrings.FString(Strings.Join(Code, "\\n"), this.Name)
    End If

    Immediate.WriteLine PStrings.FString(CLI.Lang.GetValue("messages/cstrAdded"), this.FullPath)
End Sub

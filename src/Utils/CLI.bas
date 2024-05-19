Attribute VB_Name = "CLI"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Property Get Commands() As Variant
    Commands = Array( _
        "export", _
        "help", _
        "install" _
    )
End Property

#If DEV Then
  Public Property Get Aliases() As Dictionary
#Else
  Public Property Get Aliases() As Object
#End If
  #If DEV Then
    Dim Buffer As Dictionary: Set Buffer = NewDictionary(vbTextCompare)
  #Else
    Dim Buffer As Object: Set Buffer = NewDictionary(vbTextCompare)
  #End If

    Buffer("i") = "install"
    Buffer("add") = "install"

    Buffer("exp") = "export"
    Buffer("save") = "export"

    Set Aliases = Buffer
End Property

Public Function ParseCommand(ByRef Args As Variant) As ICommand
    If Not IsArray(Args) Then
        Set ParseCommand = NewHelpCommand(Empty)
        Exit Function
    ElseIf PArray.IncludesAny(Args, "-h") Or PArray.IncludesAny(Args, "--help") Then
        Set ParseCommand = NewHelpCommand(Args)
        Exit Function
    End If

    Dim Command As String: Command = CLI.FindCommand(Args(0))
    If Strings.Len(Command) = 0 Then
        Immediate.WriteLine "Unknown command " & Args(0)
        End
    End If

    Set ParseCommand = Application.Run(FString("New{0}Command", Command), Args)
End Function

Public Function FindCommand(ByVal Name As String) As String
    If PArray.IncludesAny(Commands, Name) Then
        FindCommand = Name
        Exit Function
    End If

    If Aliases.Exists(Name) Then
        FindCommand = Aliases(Name)
        Exit Function
    End If
End Function

Public Function GetPPMProjectsPath() As String
    With NewFileSystemObject()
        Dim PPMPath As String
        PPMPath = .BuildPath(Interaction.Environ("LOCALAPPDATA"), "ppm")
        Dim PPMProjectsPath As String
        PPMProjectsPath = .BuildPath(PPMPath, "projects")
        If Not .FolderExists(PPMProjectsPath) Then CreateFoldersRecoursive PPMProjectsPath
        GetPPMProjectsPath = PPMProjectsPath
    End With
End Function

Public Function GetProjectPath() As String
    Dim PPMProjectsPath As String: PPMProjectsPath = GetPPMProjectsPath()
    With NewFileSystemObject()
        Dim ProjectName As String
        ProjectName = .GetFileName(SelectedProject.Path)

        Dim ProjectTimeStamp As String
        ProjectTimeStamp = Strings.Format( _
            .GetFile(SelectedProject.Path).DateCreated, "ddmmyyyy_hhnnss" _
        )

        Dim FolderName As String
        FolderName = FString("{0}_{1}", ProjectName, ProjectTimeStamp)

        Dim ThisProjectPath As String
        ThisProjectPath = .BuildPath(PPMProjectsPath, FolderName)
        If Not .FolderExists(ThisProjectPath) Then FileSystem.MkDir ThisProjectPath
    End With

    GetProjectPath = ThisProjectPath
End Function

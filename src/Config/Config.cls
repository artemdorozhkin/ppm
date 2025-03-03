VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed
'@Folder "PearPMProject.src.Config"
Option Explicit

Private Type TConfig
    File As ConfigIO
End Type

Private this As TConfig

Public Property Get File() As ConfigIO
    Set File = this.File
End Property

Public Sub SetScope(ByVal Scope As ConfigScopes)
    If Scope = ConfigScopes.DefaultScope Then
        Set this.File = GetDefaultConfig()
        Exit Sub
    End If

  #If DEV Then
    Dim FSO As FileSystemObject
  #Else
    Dim FSO As Object
  #End If
    Set FSO = NewFileSystemObject()

    Dim Folder As String
    Select Case Scope
        Case ConfigScopes.ProjectScope
            Folder = SelectedProject.Folder
        Case ConfigScopes.UserScope
            Folder = FSO.BuildPath(Interaction.Environ("APPDATA"), "ppm")
        Case ConfigScopes.GlobalScode
            Folder = FSO.BuildPath(Interaction.Environ("PROGRAMDATA"), "ppm")
        Case Else
            Information.Err.Raise _
                9, _
                Information.TypeName(Me), _
                PStrings.FString(CLI.Lang.GetValue("messages/unknownConfigType"), Scope)

    End Select

    If Not FSO.FolderExists(Folder) Then
        FSO.CreateFolder Folder
    End If

    Dim FilePath As String
    FilePath = FSO.BuildPath(Folder, ".ppmrc")
    Set this.File = NewConfigIO(FilePath)
End Sub

Public Function GetValue(ByVal Key As String) As Variant
  #If DEV Then
    Dim Data As Dictionary
  #Else
    Dim Data As Object
  #End If
    Set Data = this.File.Read()

    If Not Definitions.Items().Exists(Key) Then
        GetValue = Data(Key)
        Exit Function
    End If

    Dim Definition As Definition
    Set Definition = Definitions(Key)

    If Data.Exists(Definition.Key) Then
        GetValue = Utils.ConvertToType(Data(Definition.Key), Definition.DataType)
        Exit Function
    End If
End Function

Private Function GetDefaultConfig() As ConfigIO
  #If DEV Then
    Dim FSO As FileSystemObject
  #Else
    Dim FSO As Object
  #End If
    Set FSO = NewFileSystemObject()

    If SelectedProject.IsSaved Then
        Dim FilePath As String
        FilePath = PFileSystem.BuildPath(SelectedProject.Folder, ".ppmrc")
        If FSO.FileExists(FilePath) Then
            Set GetDefaultConfig = NewConfigIO(FilePath)
            Exit Function
        End If
    Else
        FilePath = PFileSystem.BuildPath(Interaction.Environ("APPDATA"), "ppm", ".ppmrc")
        If FSO.FileExists(FilePath) Then
            Set GetDefaultConfig = NewConfigIO(FilePath)
            Exit Function
        End If

        FilePath = PFileSystem.BuildPath(Interaction.Environ("PROGRAMDATA"), "ppm", ".ppmrc")
        If FSO.FileExists(FilePath) Then
            Set GetDefaultConfig = NewConfigIO(FilePath)
            Exit Function
        End If
    End If

    ' otherwise set user config as default
    FilePath = PFileSystem.BuildPath(Interaction.Environ("APPDATA"), "ppm", ".ppmrc")
    Set GetDefaultConfig = NewConfigIO(FilePath)
End Function

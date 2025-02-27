VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "PearPMProject.src.ProjectExporter"
Option Explicit

Private Type TProjectExporter
  #If DEV Then
    Components As Dictionary
  #Else
    Components As Object
  #End If
    SaveStruct As Boolean
    Destination As String
    RewriteLastExport As Boolean
    Folders As Variant
    LastExport As Variant
End Type

Private this As TProjectExporter

#If DEV Then
  Public Property Get Components() As Dictionary
#Else
  Public Property Get Components() As Object
#End If
    Set Components = this.Components
End Property

Public Property Get SaveStruct() As Boolean
    SaveStruct = this.SaveStruct
End Property
Public Property Let SaveStruct(ByVal RHS As Boolean)
    this.SaveStruct = RHS
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property
Public Property Let Destination(ByVal RHS As String)
    this.Destination = RHS
End Property

Public Property Get RewriteLastExport() As Boolean
    RewriteLastExport = this.RewriteLastExport
End Property
Public Property Let RewriteLastExport(ByVal RHS As Boolean)
    this.RewriteLastExport = RHS
End Property

Public Sub ExportAll(ByVal Encoding As String)
    this.LastExport = GetLastExport()
    If this.RewriteLastExport Then
        If UBound(this.LastExport) > -1 Then DeleteLastExport
    End If

    If Strings.Len(this.Destination) = 0 Then
        this.Destination = SelectedProject.Folder
    End If

    If this.SaveStruct Then
        Set this.Components = ParseComponentsWithDirectory()
    Else
        Set this.Components = ParseComponents()
    End If
    
    If this.RewriteLastExport Then
        this.LastExport = SaveComponents(Encoding)
        SaveLastExport
    Else
        SaveComponents Encoding
    End If
End Sub

Public Sub ExportFolders(ByVal Encoding As String, ParamArray Folders() As Variant)
    this.LastExport = GetLastExport()
    If this.RewriteLastExport Then
        If UBound(this.LastExport) > -1 Then DeleteLastExport
    End If

    this.Folders = Folders

    If Strings.Len(this.Destination) = 0 Then
        this.Destination = SelectedProject.Folder
    End If

    If this.SaveStruct Then
        Set this.Components = ParseComponentsWithDirectory()
    Else
        Set this.Components = ParseComponents()
    End If

    If this.RewriteLastExport Then
        this.LastExport = SaveComponents(Encoding)
        SaveLastExport
    Else
        SaveComponents Encoding
    End If
End Sub

Private Function GetLastExport() As Variant
    Dim Converter As BinaryConverter
    Set Converter = New BinaryConverter
    On Error Resume Next
    Dim Decoded As String
    Decoded = Strings.StrConv(PFileSystem.ReadFile(GetLastExportPath()), vbUnicode)
    GetLastExport = Strings.Split(Decoded, vbNewLine)
    If Err Then
        GetLastExport = Array()
    End If
End Function

Private Sub SaveLastExport()
    Dim Encoded As String
    Encoded = Strings.StrConv(Strings.Join(this.LastExport, vbNewLine), vbFromUnicode)
    PFileSystem.SaveToFile GetLastExportPath(), Encoded
End Sub

Private Function GetLastExportPath() As String
    GetLastExportPath = PFileSystem.BuildPath(SelectedProject.Folder, ".ppm", ".lel")
End Function

Private Sub DeleteLastExport()
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If

    Dim Item As Variant
    For Each Item In this.LastExport
        On Error GoTo Catch
        If FSO.FolderExists(Item) Then
            FSO.GetFolder(Item).Delete Force:=True
        ElseIf FSO.FileExists(Item) Then
            FSO.GetFile(Item).Delete Force:=True
        End If
    Next
Exit Sub

Catch:
    If FSO.FolderExists(Item) Then
        Immediate.WriteLine PStrings.FString( _
            "ERR: Cannot delete folder '{0}'. #{1} {2}", _
            FSO.GetFolder(Item).Name, Information.Err.Number, Information.Err.Description _
        )
    ElseIf FSO.FileExists(Item) Then
        Immediate.WriteLine PStrings.FString( _
            "ERR: Cannot delete file '{0}'. #{1} {2}", _
            FSO.GetFolder(Item).Name, _
            Information.Err.Number, _
            Information.Err.Description _
        )
    End If
End Sub

#If DEV Then
  Private Function ParseComponents() As Dictionary
#Else
  Private Function ParseComponents() As Object
#End If
  #If DEV Then
    Dim Components As Dictionary: Set Components = NewDictionary()
    Dim Component As VBComponent
  #Else
    Dim Components As Object: Set Components = NewDictionary()
    Dim Component As Object
  #End If
    For Each Component In SelectedProject.Components
        Dim Directory As String
        Directory = NewModule(Component).Path
        If IsExportingFolder(Directory) Then
            Set Components(Component.Name) = Component
        End If
    Next

    Set ParseComponents = Components
End Function

#If DEV Then
  Private Function ParseComponentsWithDirectory() As Dictionary
#Else
  Private Function ParseComponentsWithDirectory() As Object
#End If
  #If DEV Then
    Dim Components As Dictionary: Set Components = NewDictionary()
    Dim Component As VBComponent
  #Else
    Dim Components As Object: Set Components = NewDictionary()
    Dim Component As Object
  #End If
    For Each Component In SelectedProject.Components
        Dim Directory As String
        Directory = NewModule(Component).Path
        If IsExportingFolder(Directory) Then
            Dim ComponentKey As String
            ComponentKey = PStrings.FString("{0}:{1}", Component.Name, Directory)
            Set Components(ComponentKey) = Component
        End If
    Next

    Set ParseComponentsWithDirectory = Components
End Function

Private Function IsExportingFolder(ByVal Directory As String) As Boolean
    If Information.IsEmpty(this.Folders) Then
        IsExportingFolder = True
        Exit Function
    End If

    Dim Folder As Variant
    For Each Folder In this.Folders
        If PStrings.StartsWith(Directory, Folder) Then
            IsExportingFolder = True
            Exit Function
        End If
    Next
End Function

Private Function SaveComponents(ByVal Encoding As String) As Variant
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim FileList As Collection: Set FileList = New Collection

    Dim ComponentKey As Variant
    For Each ComponentKey In this.Components
      #If DEV Then
        Dim Component As VBComponent: Set Component = this.Components(ComponentKey)
      #Else
        Dim Component As Object: Set Component = this.Components(ComponentKey)
      #End If
        Dim SaveFolderPath As String
        If Not this.SaveStruct Then
            SaveFolderPath = this.Destination
        Else
            Dim DirectoryPath As String
            DirectoryPath = Strings.Split(ComponentKey, ":")(1)
            DirectoryPath = Strings.Replace(DirectoryPath, SelectedProject.Name & ".", "", Count:=1)
            DirectoryPath = Strings.Replace(DirectoryPath, SelectedProject.Name, "", Count:=1)
            DirectoryPath = Strings.Replace(DirectoryPath, ".", Application.PathSeparator)
            
            SaveFolderPath = FSO.BuildPath(this.Destination, DirectoryPath)
            If Not FSO.FolderExists(SaveFolderPath) Then PFileSystem.CreateFolder SaveFolderPath, Recoursive:=True
        End If
        Dim FileName As String
        FileName = SelectedProject.ExportComponent(SaveFolderPath, Component)
        PFileSystem.ChangeFileEncoding FileName, Encoding

        If SaveFolderPath <> this.Destination Then FileList.Add SaveFolderPath
        FileList.Add FileName
    Next

    SaveComponents = PCollection.ToArray(FileList)
End Function

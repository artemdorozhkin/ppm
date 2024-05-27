VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.ProjectExporter"
Option Explicit

Private Type TProjectExporter
    Modules As Object
    SaveStruct As Boolean
    Destination As String
    RewriteLastExport As Boolean
    Folders As Variant
    LastExport As Variant
End Type

Private this As TProjectExporter

Public Property Get Modules() As Object
    Set Modules = this.Modules
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

    If Not this.SaveStruct Then
        Set this.Modules = ParseModules()
    Else
        Set this.Modules = ParseModulesWithDirectory()
    End If

    If this.RewriteLastExport Then
        this.LastExport = SaveModules(Encoding)
        SaveLastExport
    Else
        SaveModules Encoding
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

    If Not this.SaveStruct Then
        Set this.Modules = ParseModules()
    Else
        Set this.Modules = ParseModulesWithDirectory()
    End If

    If this.RewriteLastExport Then
        this.LastExport = SaveModules(Encoding)
        SaveLastExport
    Else
        SaveModules Encoding
    End If
End Sub

Private Function GetLastExport() As Variant
    Dim ConfigIO As ConfigIO
    Set ConfigIO = NewConfigIO(GetLastExportPath())
    Dim LastExportList As String: LastExportList = ConfigIO.Read()("LastExportList")
    GetLastExport = Strings.Split(LastExportList, ";")
End Function

Private Sub SaveLastExport()
    Dim ConfigIO As ConfigIO
    Set ConfigIO = NewConfigIO(GetLastExportPath())
    ConfigIO.CreateOrWrite "LastExportList", Strings.Join(this.LastExport, ";")
End Sub

Private Function GetLastExportPath() As String
    GetLastExportPath = NewFileSystemObject().BuildPath(Constants.ProjectPath, ".lel")
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
  Private Function ParseModules() As Dictionary
#Else
  Private Function ParseModules() As Object
#End If
  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
    Dim Module As VBComponent
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As Object
  #End If
    For Each Module In SelectedProject.Modules
        Dim Directory As String
        Directory = GetDirectory(Module.CodeModule)
        If IsExportingFolder(Directory) Then
            Set Modules(Module.Name) = Module
        End If
    Next

    Set ParseModules = Modules
End Function

#If DEV Then
  Private Function ParseModulesWithDirectory() As Dictionary
#Else
  Private Function ParseModulesWithDirectory() As Object
#End If
  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
    Dim Module As VBComponent
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As Object
  #End If
    For Each Module In SelectedProject.Modules
        Dim Directory As String
        Directory = GetDirectory(Module.CodeModule)
        If IsExportingFolder(Directory) Then
            Dim ModuleKey As String
            ModuleKey = PStrings.FString("{0}:{1}", Module.Name, Directory)
            Set Modules(ModuleKey) = Module
        End If
    Next

    Set ParseModulesWithDirectory = Modules
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

#If DEV Then
  Private Function GetDirectory(ByRef Code As CodeModule) As String
#Else
  Private Function GetDirectory(ByRef Code As Object) As String
#End If
    Dim Lines As Variant: Lines = Strings.Split(Code.Lines(1, Code.CountOfDeclarationLines), vbNewLine)
    Dim Line As Variant
    For Each Line In Lines
        Dim i As Long
        i = PStrings.GetIndexOfChar(Line, "'")
        If i = -1 Then GoTo Continue

        i = PStrings.GetIndexOfChar(Line, "@", i)
        If i = -1 Then GoTo Continue

        If Not PStrings.Contains(Strings.Mid(Line, i), "Folder") Then GoTo Continue
        i = PStrings.GetIndexOfChar(Line, """", i)
        If i = -1 Then GoTo Continue

        Dim j As Long
        j = PStrings.GetIndexOfChar(Strings.StrReverse(Line), """")
        j = Strings.Len(Line) - j + 1
        If j = i Then GoTo Continue

        GetDirectory = Strings.Mid(Line, i + 1, j - i - 1)
Continue:
    Next
End Function

Private Function SaveModules(ByVal Encoding As String) As Variant
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim FileList As Collection: Set FileList = New Collection

    Dim ModuleKey As Variant
    For Each ModuleKey In this.Modules
      #If DEV Then
        Dim Module As VBComponent: Set Module = this.Modules(ModuleKey)
      #Else
        Dim Module As Object: Set Module = this.Modules(ModuleKey)
      #End If
        Dim SaveFolderPath As String
        If Not this.SaveStruct Then
            SaveFolderPath = this.Destination
        Else
            Dim DirectoryPath As String
            DirectoryPath = Strings.Split(ModuleKey, ":")(1)
            DirectoryPath = Strings.Replace(DirectoryPath, SelectedProject.Name & ".", "", Count:=1)
            DirectoryPath = Strings.Replace(DirectoryPath, ".", Application.PathSeparator)
            
            SaveFolderPath = FSO.BuildPath(this.Destination, DirectoryPath)
            If Not FSO.FolderExists(SaveFolderPath) Then CreateFoldersRecoursive SaveFolderPath
        End If
        Dim FileName As String
        FileName = SelectedProject.ExportModule(SaveFolderPath, Module)
        PFileSystem.ChangeFileEncoding FileName, Encoding

        If SaveFolderPath <> this.Destination Then FileList.Add SaveFolderPath
        FileList.Add FileName
    Next

    SaveModules = PCollection.ToArray(FileList)
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Export"
Option Explicit

Implements ICommand

Private Type TExportCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "export"
    this.CommandInfo.Description = "Exports the selected project."
    this.CommandInfo.Params.Item("encoding") = "\\t\\tExport files with set encoding."
    this.CommandInfo.Params.Item("save-struct") = "\\tSave the RubberDuck structure when exporting a project."
    this.CommandInfo.Usage = Array("[<options>]")
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
    Dim Encoding As String: Encoding = Me.GetEncoding()

    Dim LastExport As Variant: LastExport = Me.GetLastExport()
    If UBound(LastExport) > -1 Then Me.DeleteFiles LastExport

  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
  #End If
    Dim SaveStruct As Boolean
    SaveStruct = this.Config.GetValue("save-struct")

    If SaveStruct Then
        Set Modules = Me.ParseModulesWithDirectory()
    Else
        Set Modules = Me.ParseModules()
    End If

    LastExport = Me.SaveModules(Modules, Encoding, SaveStruct)
    SaveLastExport LastExport

    ms = DateTime.Timer - ms
    Immediate.WriteLine PStrings.FString( _
        "exported {0} module{1} in {2}\\nroot folder: {3}\\n", _
        Modules.Count, IIf(Modules.Count > 1, "s", ""), ConvertTime(ms), SelectedProject.Folder _
    )
End Sub

Public Function GetLastExport() As Variant
    Dim LastExportPath As String
    LastExportPath = NewFileSystemObject().BuildPath(Constants.ProjectPath, ".lel")
    Dim ConfigIO As ConfigIO: Set ConfigIO = NewConfigIO(LastExportPath)
    Dim LastExportList As String: LastExportList = ConfigIO.Read()("LastExportList")
    GetLastExport = Strings.Split(LastExportList, ";")
End Function

Public Sub SaveLastExport(ByRef LastExport As Variant)
    Dim LastExportPath As String
    LastExportPath = NewFileSystemObject().BuildPath(Constants.ProjectPath, ".lel")
    Dim ConfigIO As ConfigIO: Set ConfigIO = NewConfigIO(LastExportPath)
    ConfigIO.CreateOrWrite "LastExportList", Strings.Join(LastExport, ";")
End Sub

Public Sub DeleteFiles(ByRef List As Variant)
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If

    Dim Item As Variant
    For Each Item In List
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
  Public Function ParseModules() As Dictionary
#Else
  Public Function ParseModules() As Object
#End If
  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
    Dim Module As VBComponent
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As Object
  #End If
    For Each Module In SelectedProject.Modules
        Set Modules(Module.Name) = Module
    Next

    Set ParseModules = Modules
End Function

#If DEV Then
  Public Function ParseModulesWithDirectory() As Dictionary
#Else
  Public Function ParseModulesWithDirectory() As Object
#End If
  #If DEV Then
    Dim Modules As Dictionary: Set Modules = NewDictionary()
    Dim Module As VBComponent
  #Else
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As Object
  #End If
    For Each Module In SelectedProject.Modules
        Dim DirectoryPath As String
        DirectoryPath = PStrings.FString("{0}:{1}", Module.Name, Me.GetDirectory(Module.CodeModule))
        Set Modules(DirectoryPath) = Module
    Next

    Set ParseModulesWithDirectory = Modules
End Function

#If DEV Then
  Public Function GetDirectory(ByRef Code As CodeModule) As String
#Else
  Public Function GetDirectory(ByRef Code As Object) As String
#End If
    Dim Lines As Variant: Lines = Strings.Split(Code.Lines(1, Code.CountOfLines), vbNewLine)
    Dim Line As Variant
    For Each Line In Lines
        Dim i As Long
        i = Me.GetIndexOfChar(Line, "'")
        If i = -1 Then GoTo Continue

        i = Me.GetIndexOfChar(Line, "@", i)
        If i = -1 Then GoTo Continue

        If Not Contains(Strings.Mid(Line, i), "Folder") Then GoTo Continue
        i = Me.GetIndexOfChar(Line, """", i)
        If i = -1 Then GoTo Continue

        Dim j As Long
        j = Me.GetIndexOfChar(Strings.StrReverse(Line), """")
        j = Strings.Len(Line) - j + 1
        If j = i Then GoTo Continue

        Dim Directory As String: Directory = Strings.Mid(Line, i + 1, j - i - 1)
        Directory = Strings.Replace(Directory, SelectedProject.Name & ".", "")
        GetDirectory = Directory
Continue:
    Next
End Function

Public Function GetIndexOfChar(ByVal Text As String, ByVal Char As String, Optional ByVal StartFrom As Long = 1) As Long
    Dim i As Long: i = 1
    Dim Current As String
    Current = Strings.Mid(Text, i, 1)
    Do While Current <> Char
        i = i + 1
        If i > Strings.Len(Text) Then Exit Do
        Current = Strings.Mid(Text, i, 1)
    Loop

    If Current = Char Then
        GetIndexOfChar = i
    Else
        GetIndexOfChar = -1
    End If
End Function

Public Function GetEncoding() As Variant
    Dim EncodingDef As Definition: Set EncodingDef = Definitions("encoding")
    If this.Tokens.IncludeDefinition(EncodingDef) Then
        GetEncoding = this.Tokens.GetTokenByDefinition(EncodingDef).Text
    Else
        GetEncoding = EncodingDef.Default
    End If
End Function

#If DEV Then
  Public Function SaveModules(ByRef Modules As Dictionary, ByVal Encoding As String, ByVal SaveStruct As Boolean) As Variant
#Else
  Public Function SaveModules(ByRef Modules As Object, ByVal Encoding As String, ByVal SaveStruct As Boolean) As Variant
#End If
  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = NewFileSystemObject()
  #End If
    Dim RootPath As String: RootPath = FSO.GetParentFolderName(SelectedProject.Path)
    Dim FileList As Collection: Set FileList = New Collection

    Dim DirectoryPath As Variant
    For Each DirectoryPath In Modules
      #If DEV Then
        Dim Module As VBComponent: Set Module = Modules(DirectoryPath)
      #Else
        Dim Module As Object: Set Module = Modules(DirectoryPath)
      #End If
        Dim SaveFolderPath As String
        If Not SaveStruct Then
            SaveFolderPath = RootPath
        Else
            DirectoryPath = Strings.Split(DirectoryPath, ":")(1)
            DirectoryPath = Strings.Replace(DirectoryPath, ".", Application.PathSeparator)
            SaveFolderPath = FSO.BuildPath(RootPath, DirectoryPath)
            If Not FSO.FolderExists(SaveFolderPath) Then CreateFoldersRecoursive SaveFolderPath
        End If
        Dim FileName As String
        FileName = SelectedProject.ExportModule(SaveFolderPath, Module)
        Me.ChangeFileEncoding FileName, Encoding

        If SaveFolderPath <> RootPath Then FileList.Add SaveFolderPath
        FileList.Add FileName
    Next

    SaveModules = PCollection.ToArray(FileList)
End Function

Public Sub ChangeFileEncoding(ByVal Path As String, ByVal Encoding As String)
  #If DEV Then
      Dim SourceStream As TextStream
  #Else
    Dim SourceStream As Object
  #End If
    Set SourceStream = NewFileSystemObject().OpenTextFile(Path)
    Dim Content As String
    If Not SourceStream.AtEndOfStream Then
        Content = SourceStream.ReadAll()
    End If
    SourceStream.Close

    SaveToFile Path, Content, Encoding
End Sub

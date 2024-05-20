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
    Tokens As Tokens
    CommandInfo As CommandInfo
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "export"
    this.CommandInfo.Description = "Exports the selected project."
    this.CommandInfo.Params = Array( _
        "encoding", _
        "save-struct" _
    )
    this.CommandInfo.Usage = Array("[<options>]")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Public Function GetLastExport() As Variant
    Dim ConfigIO As ConfigIO: Set ConfigIO = Configs.GetCurrentProjectConfig()
    Dim LastExport As Object: Set LastExport = ConfigIO.ReadSection("LastExport")
    GetLastExport = Strings.Split(LastExport("List"), ";")
End Function

Public Sub SaveLastExport(ByRef LastExport As Variant)
    Dim ConfigIO As ConfigIO: Set ConfigIO = Configs.GetCurrentProjectConfig()
    ConfigIO.WriteData "LastExport", "List", Strings.Join(LastExport, ";")
End Sub

Public Sub DeleteFiles(ByRef List As Variant)
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()

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
        Immediate.WriteLine FString( _
            "ERR: Cannot delete folder '{0}'. #{1} {2}", _
            FSO.GetFolder(Item).Name, Err.Number, Err.Description _
        )
    ElseIf FSO.FileExists(Item) Then
        Immediate.WriteLine FString( _
            "ERR: Cannot delete file '{0}'. #{1} {2}", _
            FSO.GetFolder(Item).Name, Err.Number, Err.Description _
        )
    End If
End Sub

Public Function ParseModules() As Object
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As VBComponent
    For Each Module In SelectedProject.Modules
        Set Modules(Module.Name) = Module
    Next

    Set ParseModules = Modules
End Function

Public Function ParseModulesWithDirectory() As Object
    Dim Modules As Object: Set Modules = NewDictionary()
    Dim Module As VBComponent
    For Each Module In SelectedProject.Modules
        Dim DirectoryPath As String
        DirectoryPath = FString("{0}:{1}", Module.Name, Me.GetDirectory(Module.CodeModule))
        Set Modules(DirectoryPath) = Module
    Next

    Set ParseModulesWithDirectory = Modules
End Function

Public Function GetDirectory(ByRef Code As CodeModule) As String
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
    Dim Encoding As String: Encoding = "UTF-8"

    Dim EncodingTokenIndex As Long
    EncodingTokenIndex = this.Tokens.GetTokenIndex("encoding", OptionItem)
    If EncodingTokenIndex = -1 Then EncodingTokenIndex = this.Tokens.GetTokenIndex("e", ShortOptionItem)

    If EncodingTokenIndex = -1 Then
        GetEncoding = Encoding
    Else
        GetEncoding = this.Tokens.MatchToken( _
            EncodingTokenIndex + 1, Identifier, "Expected encoding type after ""-e|--encoding"" option." _
        ).Text
    End If

    GetEncoding = Encoding
End Function

Public Function SaveModules(ByRef Modules As Object, ByVal Encoding As String, ByVal SaveStruct As Boolean) As Variant
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
    Dim RootPath As String: RootPath = FSO.GetParentFolderName(SelectedProject.Path)
    Dim FileList As Collection: Set FileList = New Collection

    Dim DirectoryPath As Variant
    For Each DirectoryPath In Modules
        Dim Module As VBComponent: Set Module = Modules(DirectoryPath)
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
    Dim SourceStream As Object
    Set SourceStream = NewFileSystemObject().OpenTextFile(Path)
    Dim Content As String
    If Not SourceStream.AtEndOfStream Then
        Content = SourceStream.ReadAll()
    End If
    SourceStream.Close

    SaveToFile Path, Content, Encoding
End Sub

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    Dim ms As Double: ms = DateTime.Timer
    Dim Encoding As String: Encoding = Me.GetEncoding()

    Dim LastExport As Variant: LastExport = Me.GetLastExport()
    If UBound(LastExport) > -1 Then Me.DeleteFiles LastExport

    Dim Modules As Dictionary: Set Modules = NewDictionary()
    Dim SaveStruct As Boolean: SaveStruct = this.Tokens.IncludeToken("s", ShortOptionItem) Or _
                                            this.Tokens.IncludeToken("save-struct", OptionItem)
    If SaveStruct Then
        Set Modules = Me.ParseModulesWithDirectory()
    Else
        Set Modules = Me.ParseModules()
    End If

    LastExport = Me.SaveModules(Modules, Encoding, SaveStruct)
    SaveLastExport LastExport

    ms = DateTime.Timer - ms
    Immediate.WriteLine FString( _
        "exported {0} module{1} in {2}\\nroot folder: {3}\\n", _
        Modules.Count, IIf(Modules.Count > 1, "s", ""), ConvertTime(ms), SelectedProject.Folder _
    )
End Sub



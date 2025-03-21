Attribute VB_Name = "PFileSystem"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Public Sub ChangeFileEncoding(ByVal Path As String, ByVal Encoding As String)
  #If DEV Then
    Dim SourceStream As TextStream
  #Else
    Dim SourceStream As Object
  #End If
    Set SourceStream = GetFileSystemObject().OpenTextFile(Path)
    Dim Content As String
    If Not SourceStream.AtEndOfStream Then
        Content = SourceStream.ReadAll()
    End If
    SourceStream.Close

    SaveToFile Path, Content, Encoding
End Sub

Public Function ReadFile(ByVal Path As String, Optional ByVal Encoding As String = "UTF-8") As String
    With NewStream()
        .Charset = Encoding
        .Type = 2 'adTypeText
        .Open
        .LoadFromFile Path
        ReadFile = .ReadText()
        .Close
    End With
End Function

Public Sub SaveToFile(ByVal Path As String, ByVal Content As String, Optional ByVal Encoding As String = "UTF-8")
    Dim IsUTF8 As Boolean: IsUTF8 = PStrings.IsEqual(Encoding, "utf-8")

  #If DEV Then
    Dim EncodingStream As Stream: Set EncodingStream = NewStream()
  #Else
    Dim EncodingStream As Object: Set EncodingStream = NewStream()
  #End If
    EncodingStream.Mode = 3 'adModeReadWrite
    EncodingStream.Charset = Encoding
    EncodingStream.Open
    EncodingStream.WriteText Content
    EncodingStream.Position = Interaction.IIf(IsUTF8, 3, 0) 'skip BOM

  #If DEV Then
    Dim BinaryStream As Stream: Set BinaryStream = NewStream()
  #Else
    Dim BinaryStream As Object: Set BinaryStream = NewStream()
  #End If
    BinaryStream.Mode = 3 'adModeReadWrite
    BinaryStream.Type = 1 'adTypeBinary
    BinaryStream.Open

    EncodingStream.CopyTo BinaryStream
    EncodingStream.Close

    With GetFileSystemObject()
        If Not .FileExists(Path) Then
            PFileSystem.CreateFolder .GetParentFolderName(Path), Recoursive:=True
            BinaryStream.SaveToFile Path, 1 'adSaveCreateNotExist
        Else
            BinaryStream.SaveToFile Path, 2 'adSaveCreateOverWrite
        End If
    End With
    BinaryStream.Close
End Sub

Public Sub CreateFolder(ByVal Path As String, Optional ByVal Recoursive As Boolean = False)
    If Not Recoursive Then
        FileSystem.MkDir Path
        Exit Sub
    End If

  #If DEV Then
    Dim FSO As FileSystemObject: Set FSO = NewFileSystemObject()
  #Else
    Dim FSO As Object: Set FSO = GetFileSystemObject()
  #End If
    Dim Parts As Variant: Parts = Strings.Split(Path, Application.PathSeparator)
    Dim Part As Variant
    For Each Part In Parts
        Dim Current As String
        If Strings.Len(Current) = 0 Then
            Current = Part
        Else
            Current = PStrings.FString("{0}{1}{2}", Current, Application.PathSeparator, Part)
        End If

        If Not FSO.FolderExists(Current) Then
            FSO.CreateFolder Current
        End If
    Next
End Sub

Public Function GetFileNameWithoutExt(ByVal Path As String) As String
    Dim FileName As String: FileName = GetFileName(Path)
    Dim Ext As String: Ext = GetFileExt(Path)
    GetFileNameWithoutExt = Strings.Left(FileName, Strings.Len(FileName) - Strings.Len(Ext))
End Function

Public Function GetFileName(ByVal Path As String) As String
    GetFileName = Strings.Mid(Path, Strings.InStrRev(Path, Application.PathSeparator) + 1)
End Function

Public Function GetFileExt(ByVal Path As String) As String
    Dim FileName As String: FileName = GetFileName(Path)
    Dim DotPosition As Long: DotPosition = Strings.InStrRev(FileName, ".")
    If DotPosition = 0 Then Exit Function
    GetFileExt = Strings.Mid(FileName, DotPosition)
End Function

Public Function BuildPath(ParamArray Parts() As Variant)
  #If DEV Then
    Dim FSO As FileSystemObject
  #Else
    Dim FSO As Object
  #End If
    Set FSO = GetFileSystemObject()

    Dim FullPath As String
    Dim Part As Variant
    For Each Part In Parts
        FullPath = FSO.BuildPath(FullPath, Part)
    Next

    BuildPath = FullPath
End Function

Attribute VB_Name = "PQuickConstructors"
'@Folder "PearPMProject.src.Utils"
Option Explicit

#If DEV Then
  Public Function NewStream() As Stream
#Else
  Public Function NewStream() As Object
#End If
    Set NewStream = CreateObject("ADODB.Stream")
End Function

#If DEV Then
  Public Function NewDictionary( _
    Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
  ) As Dictionary
#Else
  Public Function NewDictionary( _
    Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
  ) As Object
#End If
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = Compare
End Function

#If DEV Then
  Public Function GetFileSystemObject() As FileSystemObject
#Else
  Public Function GetFileSystemObject() As Object
#End If
    Static FSO As Object
    If FSO Is Nothing Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set GetFileSystemObject = FSO
End Function

#If DEV Then
  Public Function NewFolder(ByVal Path As String) As Scripting.Folder
#Else
  Public Function NewFolder(ByVal Path As String) As Object
#End If
    With GetFileSystemObject()
        If (FileSystem.GetAttr(Path) And vbDirectory) = vbDirectory Then
            Set NewFolder = .GetFolder(Path)
        ElseIf .FileExists(Path) Then
            Set NewFolder = .GetFile(Path).ParentFolder
        End If
    End With
End Function

#If DEV Then
  Public Function NewShell() As Shell32.Shell
#Else
  Public Function NewShell() As Object
#End If
    Set NewShell = CreateObject("Shell.Application")
End Function

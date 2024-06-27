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
  Public Function NewFileSystemObject() As FileSystemObject
#Else
  Public Function NewFileSystemObject() As Object
#End If
    Set NewFileSystemObject = CreateObject("Scripting.FileSystemObject")
End Function

#If DEV Then
  Public Function NewFolder(ByVal Path As String) As Folder
#Else
  Public Function NewFolder(ByVal Path As String) As Object
#End If
    With NewFileSystemObject()
        If (FileSystem.GetAttr(Path) And vbDirectory) = vbDirectory Then
            Set NewFolder = .GetFolder(Path)
        ElseIf .FileExists(Path) Then
            Set NewFolder = .GetFile(Path).ParentFolder
        End If
    End With
End Function

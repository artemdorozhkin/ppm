VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Utils"
Option Explicit

Private Type TBinaryConverter
  #If DEV Then
    Stream As ADODB.Stream
  #Else
    Stream As Object
  #End If
End Type

Private this As TBinaryConverter

Private Sub Class_Initialize()
    Set this.Stream = NewStream()
End Sub

Public Function TextToBytes(ByVal Text As String) As Variant
    With NewStream()
        .Open
        .Type = 2 ' text
        .Charset = "_autodetect"
        .WriteText Text
        .Position = 0
        .Type = 1
        TextToBytes = .Read()
        .Close
    End With
End Function

Public Function FileToBytes(ByVal Path As String) As Byte()
    With this.Stream
        .Type = 1 'binary
        .Open
        .LoadFromFile Path
        .Position = 0
        FileToBytes = .Read()
        .Close
    End With
End Function

Public Sub ToFile(ByRef FileBytes() As Byte, ByVal Path As String)
    With this.Stream
        .Type = 1 ' Binary type
        .Open
        .Write FileBytes

        .SaveToFile Path, 2 ' adSaveCreateOverWrite
        .Close
    End With
End Sub

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
    Stream As Object
End Type

Private this As TBinaryConverter

Private Sub Class_Initialize()
    Set this.Stream = CreateObject("ADODB.Stream")
End Sub

Public Function FromFile(ByVal Path As String) As String
    Dim fileNum As Integer: fileNum = FreeFile

    Open Path For Binary Access Read As #fileNum
    Dim FileData() As Byte
    ReDim FileData(1 To LOF(fileNum))
    Get #fileNum, , FileData
    Close #fileNum

    FromFile = Strings.StrConv(FileData, vbUnicode)
End Function

Public Sub ToFile(ByRef Bytes() As Byte, ByVal Path As String)
    With this.Stream
        .Open
        .Type = 1
        .Write Bytes
        .SaveToFile Path, 2 ' 1 = no overwrite, 2 = overwrite
        .Close
    End With
End Sub

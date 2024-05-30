VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.HTTP"
Option Explicit

Private Type TFormData
    Converter As BinaryConverter
End Type

Private this As TFormData

Public Function Create(ByVal Boundary As String, ByVal Path As String, ByVal Pack As String) As Byte()
    Set this.Converter = New BinaryConverter

    Dim FileData() As Byte
    FileData = this.Converter.FileToBytes(Path)

    With NewStream()
        Dim FormData As String
        FormData = "--" & Boundary & vbCrLf
        FormData = FormData & "Content-Disposition: form-data; name=""package"";"
        FormData = FormData & " filename=""" & FileSystem.Dir(Path) & """" & vbCrLf
        FormData = FormData & "Content-Type: application/zip" & vbCrLf & vbCrLf
        .Type = 1 'binary
        .Open
        .Write this.Converter.TextToBytes(FormData)
        .Write FileData
        .Write this.Converter.TextToBytes(vbCrLf & "--" & Boundary & vbCrLf)

        FormData = "Content-Disposition: form-data; name=""packInfo""" & vbCrLf & vbCrLf
        .Write this.Converter.TextToBytes(FormData)
        .Write this.Converter.TextToBytes(Strings.Replace(Pack, vbNewLine, ""))
        .Write this.Converter.TextToBytes(vbCrLf & "--" & Boundary & "--" & vbCrLf)
        .Position = 0

        Create = .Read()
        .Close
    End With
End Function

Public Function GenerateBoundary(Optional ByVal Seed As Long = 1) As String
    Dim CurrentDate As String
    CurrentDate = Strings.Format(DateTime.Now, "YYYYMMDDhhnnss")
    Dim BoundaryCode As String
    BoundaryCode = Conversion.Hex(CurrentDate)

    GenerateBoundary = "WebKitFormBoundary" & BoundaryCode
End Function

Function StringToBytes(str As String) As Byte()
    With CreateObject("ADODB.Stream")
        .Type = 2 'adTypeText
        .Charset = "iso-8859-1"
        .Open
        .WriteText str
        .Position = 0
        .Type = 1 'adTypeBinary
        .Position = 3 ' Skip BOM
        StringToBytes = .Read
        .Close
    End With
End Function

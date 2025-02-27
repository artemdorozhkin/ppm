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
  #If DEV Then
    Stream As Stream
    Converter As BinaryConverter
  #Else
    Stream As Object
    Converter As Object
  #End If
    Boundary As String
    ContentDisposition As String
End Type

Private this As TFormData

Public Property Get Boundary() As String
    Boundary = this.Boundary
End Property

Public Function Create() As FormData
    this.Boundary = GenerateBoundary()
    Set this.Stream = NewStream()
    Set this.Converter = New BinaryConverter
    this.ContentDisposition = "Content-Disposition: form-data; name=""{0}"";"
    Set Create = Me
End Function

Public Function AddFile(ByVal Path As String, ByVal FormDataName As String, ByVal ContentType As String) As FormData
    With this.Stream
        Dim FormData As String
        FormData = "--" & this.Boundary & vbCrLf
        FormData = FormData & PStrings.FString(this.ContentDisposition, FormDataName)
        FormData = FormData & " filename=""" & FileSystem.Dir(Path) & """" & vbCrLf
        FormData = FormData & "Content-Type: " & ContentType & vbCrLf & vbCrLf
        .Type = 1 'binary
        .Open
        .Write this.Converter.TextToBytes(FormData)
        .Write this.Converter.FileToBytes(Path)
    End With
    Set AddFile = Me
End Function

Public Function AddText(ByVal Text As String, ByVal FormDataName As String) As FormData
    With this.Stream
        .Write this.Converter.TextToBytes(vbCrLf & "--" & this.Boundary & vbCrLf)
        Dim FormData As String
        FormData = PStrings.FString(this.ContentDisposition, FormDataName) & vbCrLf & vbCrLf
        .Write this.Converter.TextToBytes(FormData)
        .Write this.Converter.TextToBytes(Strings.Replace(Text, vbNewLine, ""))
    End With
    Set AddText = Me
End Function

Public Function ToBytes() As Byte()
    With this.Stream
        .Write this.Converter.TextToBytes(vbCrLf & "--" & this.Boundary & "--" & vbCrLf)
        .Position = 0
        ToBytes = .Read()
        .Close
    End With
End Function

Private Function GenerateBoundary() As String
    Dim CurrentDate As String
    CurrentDate = Strings.Format(DateTime.Now, "YYYYMMDDhhnnss")
    Dim BoundaryCode As String
    BoundaryCode = Conversion.Hex(CurrentDate)

    GenerateBoundary = "WebKitFormBoundary" & BoundaryCode
End Function

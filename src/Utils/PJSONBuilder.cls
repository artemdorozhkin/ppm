VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJSONBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PearPMProject.src.Utils")
Option Explicit

Private Type TPJSONBuilder
    JSON As Object
End Type

Private this As TPJSONBuilder

Private Sub Class_Initialize()
    Set this.JSON = NewDictionary()
End Sub

Public Function AddObject(ParamArray KeyValuePairs() As Variant) As PJSONBuilder
    Dim i As Long
    For i = 0 To UBound(KeyValuePairs) Step 2
        this.JSON.Add KeyValuePairs(i), KeyValuePairs(i + 1)
    Next
    Set AddObject = Me
End Function

Public Function AddArray(ParamArray Values() As Variant) As Collection
    Dim Buffer As Collection: Set Buffer = New Collection
    Dim Value As Variant
    For Each Value In Values
        Buffer.Add Value
    Next

    Set AddArray = Buffer
End Function

Public Function ToString(Optional ByVal Indent As Integer = 2) As String
    ToString = PJSON.Stringify(this.JSON, Indent)
End Function

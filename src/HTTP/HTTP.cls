VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.HTTP"
Option Explicit

Private Type THTTP
  #If DEV Then
    HTTP As MSXML2.XMLHTTP60
  #Else
    HTTP As Object
  #End If
End Type

Private this As THTTP

Private Sub Class_Initialize()
    Set this.HTTP = CreateObject("MSXML2.XMLHTTP")
End Sub

Public Function GetRequest(ByVal URL As String) As TResponse
    this.HTTP.Open "GET", URL, False
    this.HTTP.Send

    GetRequest.Text = this.HTTP.ResponseText
    GetRequest.Body = this.HTTP.responseBody
    GetRequest.Code = this.HTTP.Status
End Function

#If DEV Then
  Public Function PostRequest( _
    ByVal URL As String, _
    ByRef Headers As Dictionary, _
    Optional ByVal Body As Variant _
  ) As TResponse
#Else
  Public Function PostRequest( _
    ByVal URL As String, _
    ByRef Headers As Object, _
    Optional ByVal Body As Variant _
  ) As TResponse
#End If
    this.HTTP.Open "POST", URL, False
    Dim Title As Variant
    For Each Title In Headers
        this.HTTP.SetRequestHeader Title, Headers(Title)
    Next
    this.HTTP.Send Body

    PostRequest.Text = this.HTTP.ResponseText
    PostRequest.Body = this.HTTP.responseBody
    PostRequest.Code = this.HTTP.Status
End Function

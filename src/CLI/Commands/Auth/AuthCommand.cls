VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AuthCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Commands.Auth"
Option Explicit

Implements ICommand

Private Type TAuthCommand
    CommandInfo As CommandInfo
    Tokens As Tokens
End Type

Private this As TAuthCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/auth"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "auth"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("registry") = CLI.Lang.GetValue("param", "name=registry")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Tokens(ByVal RHS As Tokens)
    Set this.Tokens = RHS
End Property

Private Property Get ICommand_CommandInfo() As CommandInfo
    Set ICommand_CommandInfo = this.CommandInfo
End Property

Private Sub ICommand_Exec()
    If Not this.Tokens.IncludeTokenKind(TokenKind.Identifier) Then
        ppm "auth --help"
        Exit Sub
    End If

    Dim EmailToken As SyntaxToken
    Set EmailToken = this.Tokens.PopFirstTokenKind(TokenKind.Identifier)
    Dim Email As String
    Email = EmailToken.Text

    Dim JSONBody As String
    JSONBody = "{""email"": """ & Email & """}"
    Dim Body() As Byte
    Body = Strings.StrConv(JSONBody, VbStrConv.vbFromUnicode)

    Dim Headers As Object
    Set Headers = NewDictionary()
    Headers("Content-Type") = "application/json"
    Headers("Content-Length") = Strings.Len(JSONBody)

    Immediate.WriteLine CLI.Lang.GetValue("messages/waitingAnswer")

    Dim Request As HTTP
    Set Request = New HTTP
    Dim Registry As String
    Registry = Utils.GetTokenOrConfigValue("registry", this.Tokens)
    If Strings.Len(Registry) = 0 Then
        Registry = Definitions("registry").Default
    End If
    Dim AuthURL As String
    AuthURL = Utils.ResolveUrl(Registry, "/auth")

    Dim Response As TResponse
    Response = Request.PostRequest(AuthURL, Headers, "{""email"": """ & Email & """}")

    If Response.Code = 400 Then
        Dim Error As Object
        Set Error = PJSON.Parse(Response.Text)
        Immediate.WriteLine CLI.Lang.GetValue("messages/errorOccurred", Error("error"))
        Exit Sub
    ElseIf Response.Code <> HTTPCodes.CREATED_201 Then
        Immediate.WriteLine CLI.Lang.GetValue("messages/somethingWentWrong")
        Exit Sub
    End If

    Dim JSON As Object
    Set JSON = PJSON.Parse(Response.Text)

    Config.ReadScope ConfigScopes.UserScope
    Config.SetValue "email", Email
    Config.SetValue "api_key", JSON("api_key")
    Config.Save ConfigScopes.UserScope

    Immediate.WriteLine PStrings.FString( _
        CLI.Lang.GetValue("messages/success"), Email, JSON("api_key") _
    )
End Sub

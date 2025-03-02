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

Private Type TExportCommand
    CommandInfo As CommandInfo
    Config As Config
    Tokens As Tokens
End Type

Private this As TExportCommand

Private Sub Class_Initialize()
    CLI.Lang.SetBlock "/root/ppmCommands/auth"

    Set this.CommandInfo = New CommandInfo
    this.CommandInfo.Name = "auth"
    this.CommandInfo.Description = CLI.Lang.GetValue("description")
    this.CommandInfo.Params.Item("registry") = CLI.Lang.GetValue("param", "name=registry")
    this.CommandInfo.Usage = CLI.Lang.GetValues("usage")
End Sub

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

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

    Immediate.WriteLine "Waiting answer from server..."

    Dim Request As HTTP
    Set Request = New HTTP
    Dim RegistryURL As String
    RegistryURL = GetFirstValueFrom("registry", this.Tokens, this.Config)
    If Strings.Len(RegistryURL) = 0 Then
        RegistryURL = Definitions("registry").Default
    End If
    Dim AuthURL As String
    AuthURL = Utils.ResolveUrl(RegistryURL, "/auth")

    Dim Response As TResponse
    Response = Request.PostRequest(AuthURL, Headers, "{""email"": """ & Email & """}")

    If Response.Code = 400 Then
        Dim Error As Object
        Set Error = PJSON.Parse(Response.Text)
        Immediate.WriteLine Error("error")
        Exit Sub
    ElseIf Response.Code <> HTTPCodes.CREATED_201 Then
        Immediate.WriteLine "internal server error!"
        Exit Sub
    End If

    Dim JSON As Object
    Set JSON = PJSON.Parse(Response.Text)

    Dim Config As Config
    Set Config = NewConfig(ConfigScopes.UserScope)
    Config.File.CreateOrWrite "email", Email
    Config.File.CreateOrWrite "api_key", JSON("api_key")

    Immediate.WriteLine FString("Your api-key for email '{0}': {1}", Email, JSON("api_key"))
End Sub

Public Function GetFirstValueFrom(ByVal DefName As String, ByRef Tokens As Tokens, ByRef Config As Config) As Variant
    If Not Definitions.Items().Exists(DefName) Then
        Err.Raise 9, "Utils", "Can't find definition: " & DefName
    End If

    Dim Def As Definition
    Set Def = Definitions(DefName)

    If Tokens.IncludeDefinition(Def) Then
        Dim Token As SyntaxToken
        Set Token = Tokens.PopTokenByDefinition(Def)
        If IsFalse(Token) Then
            GetFirstValueFrom = True
        Else
            GetFirstValueFrom = ConvertToType(Token.Text, Def.DataType)
        End If
    Else
        Dim Value As Variant
        Value = GetFirstTrue( _
            Config.GetValue(DefName), _
            Def.Default _
        )
        If Not Information.IsNull(Value) Then
            GetFirstValueFrom = ConvertToType(Value, Def.DataType)
        Else
            GetFirstValueFrom = ConvertToType(Def.Default, Def.DataType)
        End If
    End If
End Function

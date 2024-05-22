VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Lexer"
Option Explicit

Private Type TTokens
    Tokens As Collection
    CurrentEnum As IUnknown
End Type

Private this As TTokens

Private Sub Class_Initialize()
    Set this.Tokens = New Collection
End Sub

Public Property Get Count() As Long
    Count = this.Tokens.Count
End Property

'@DefaultMember
Public Property Get Token(ByVal Index As Long) As SyntaxToken
Attribute Token.VB_UserMemId = 0
    Set Token = this.Tokens(Index)
End Property

'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set this.CurrentEnum = this.Tokens.[_NewEnum]
    Set NewEnum = this.CurrentEnum
End Property

Public Sub Add(ByRef Token As SyntaxToken)
    this.Tokens.Add Token
End Sub

Public Function GetToken(ByVal Text As String, ByVal Kind As TokenKind) As SyntaxToken
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind = Kind And IsEqual(Token.Text, Text) Then
            Set GetToken = Token
            Exit Function
        End If
    Next
End Function

Public Function GetTokenIndex(ByVal Text As String, ByVal Kind As TokenKind) As Long
    Dim i As Long
    For i = 1 To this.Tokens.Count
        Dim Token As SyntaxToken: Set Token = this.Tokens(i)
        If Token.Kind = Kind And IsEqual(Token.Text, Text) Then
            GetTokenIndex = i
            Exit Function
        End If
    Next

    GetTokenIndex = -1
End Function

Public Function GetNextTokenAfter(ByVal Text As String, ByVal Kind As TokenKind) As SyntaxToken
    Dim i As Long: i = Me.GetTokenIndex(Text, Kind)
    If i = -1 Then Exit Function
    If i + 1 > this.Tokens.Count Then Exit Function
    Set GetNextTokenAfter = this.Tokens(i + 1)
End Function

Public Function GetFirstTokenKind(ByVal Kind As TokenKind) As SyntaxToken
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind = Kind Then
            Set GetFirstTokenKind = Token
            Exit Function
        End If
    Next
End Function

Public Function GetTokensKind(ByVal Kind As TokenKind) As Variant
    Dim Buffer As Collection: Set Buffer = New Collection

    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind = Kind Then
            Buffer.Add Token
        End If
    Next

    GetTokensKind = PCollection.ToArray(Buffer)
End Function

Public Function IncludeDefinition(ByRef Definition As Definition) As Boolean
    IncludeDefinition = Me.IncludeToken(Definition.Key, TokenKind.OptionItem) Or _
                        Me.IncludeToken(Definition.Short, TokenKind.ShortOptionItem)
End Function

Public Function GetTokenByDefinition(ByRef Definition As Definition) As SyntaxToken
    If Me.IncludeToken(Definition.Key, TokenKind.OptionItem) Then
        Set GetTokenByDefinition = Me.GetNextTokenAfter(Definition.Key, TokenKind.OptionItem)
    ElseIf Me.IncludeToken(Definition.Short, TokenKind.ShortOptionItem) Then
        Set GetTokenByDefinition = Me.GetNextTokenAfter(Definition.Short, TokenKind.ShortOptionItem)
    End If
End Function

Public Function IncludeToken(ByVal Text As String, ByVal Kind As TokenKind) As Boolean
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If IsEqual(Token.Text, Text) And Token.Kind = Kind Then
            IncludeToken = True
            Exit Function
        End If
    Next
End Function

Public Function IncludeTokenText(ByVal Text As String) As Boolean
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If IsEqual(Token.Text, Text) Then
            IncludeTokenText = True
            Exit Function
        End If
    Next
End Function

Public Function IncludeTokenKind(ByVal Kind As TokenKind) As Boolean
    Dim Token As SyntaxToken
    For Each Token In this.Tokens
        If Token.Kind = Kind Then
            IncludeTokenKind = True
            Exit Function
        End If
    Next
End Function

Public Function MatchToken(ByVal Index As Long, ByVal Kind As TokenKind, ByVal ErrMessage As String) As SyntaxToken
    If this.Tokens.Count < Index Then
        Information.Err.Raise 1, Information.TypeName(Me), ErrMessage
    ElseIf this.Tokens(Index).Kind = Kind Then
        Set MatchToken = this.Tokens(Index)
    Else
        Information.Err.Raise 1, Information.TypeName(Me), ErrMessage
    End If
End Function

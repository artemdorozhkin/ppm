VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.CLI.Lexer"
Option Explicit

Private Type TLexer
    Position As Long
    Start As Long
    Kind As TokenKind
    Text As String
    Tokens As Tokens
End Type

Private this As TLexer

Public Property Let Text(ByVal RHS As String)
    this.Text = RHS
End Property

Public Function Lex() As Tokens
    this.Position = 1
    Set this.Tokens = New Tokens

    Do While this.Position <= Strings.Len(this.Text)
        this.Start = this.Position
        Select Case Current
            Case "-"
                ReadOptionItem

            Case " "
                ReadWhiteSpace
                GoTo Continue

            Case "="
                this.Position = this.Position + 1
                GoTo Continue

            Case Else
                ReadIdentifier
        End Select

        Dim Length As Long: Length = this.Position - this.Start
        Dim Text As String: Text = Strings.Mid(this.Text, this.Start, Length)
        this.Tokens.Add NewSyntaxToken(this.Kind, Text)
Continue:
    Loop

    Set Lex = this.Tokens
End Function

Private Function Peek(ByVal Offset As Long) As Variant
    Dim Index As Long: Index = this.Position + Offset
    If Index > Strings.Len(this.Text) Then
        Peek = Null
        Exit Function
    End If
    Peek = Strings.Mid(this.Text, Index, 1)
End Function

Private Function Current() As Variant
    Current = Peek(0)
End Function

Private Sub ReadWhiteSpace()
    Do While IsWhiteSpace(Current)
        this.Position = this.Position + 1
        If Information.IsNull(Current) Then Exit Do
    Loop
End Sub

Private Sub ReadOptionItem()
    this.Position = this.Position + 1
    this.Start = this.Position
    this.Kind = ShortOptionItem
    If Current = "-" Then
        this.Position = this.Position + 1
        this.Start = this.Position
        this.Kind = OptionItem
    End If

    Do While Current Like "[A-Za-zА-Яа-яЁё0-9-_]"
        this.Position = this.Position + 1
        If Information.IsNull(Current) Then Exit Do
    Loop
End Sub

Private Sub ReadIdentifier()
    Do While Current Like "[A-Za-zА-Яа-яЁё0-9-_@.]"
        this.Position = this.Position + 1
        If Information.IsNull(Current) Then Exit Do
    Loop

    Dim Length As Long: Length = this.Position - this.Start
    Dim Text As String: Text = Strings.Mid(this.Text, this.Start, Length)
    If PArray.IncludesAny(CLI.Commands, Text) Or _
       CLI.Aliases.Exists(Text) Then
        this.Kind = Command
    Else
        this.Kind = Identifier
    End If
End Sub



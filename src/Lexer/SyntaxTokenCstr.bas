Attribute VB_Name = "SyntaxTokenCstr"
'@Folder "PearPMProject.src.Lexer"
Option Explicit

Public Function NewSyntaxToken(ByVal Kind As TokenKind, ByVal Text As String) As SyntaxToken
    Set NewSyntaxToken = New SyntaxToken
    NewSyntaxToken.Kind = Kind
    NewSyntaxToken.Text = Text
End Function

Attribute VB_Name = "LexerCstr"
'@Folder "PearPMProject.src.Lexer"
Option Explicit

Public Function NewLexer(ByVal Text As String) As Lexer
    Set NewLexer = New Lexer
    NewLexer.Text = Text
End Function

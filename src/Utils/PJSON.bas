Attribute VB_Name = "PJSON"
'@Folder "PearPMProject.src.Utils"
Option Explicit

Private Type TPJSON
    JSON As String
    Position As Long
End Type

Private this As TPJSON

Public Function Parse(ByVal JSONString As String) As Object
    this.JSON = JSONString
    this.Position = 1
    Set Parse = ParseValue()
End Function

#If DEV Then
  Public Function Stringify(ByVal JSONObject As Dictionary, Optional ByVal Indent As Integer = 0)
#Else
  Public Function Stringify(ByVal JSONObject As Object, Optional ByVal Indent As Integer = 0)
#End If
    Stringify = SerializeValue(JSONObject, Indent, 1)
End Function

Private Function SerializeValue(ByVal Data As Variant, ByVal Indent As Integer, ByVal Level As Integer) As String
    Dim JSONString As String
    Select Case Information.VarType(Data)
        Case VbVarType.vbString
            JSONString = """" & Data & """"

        Case VbVarType.vbBoolean
            JSONString = IIf(Data, "true", "false")

        Case VbVarType.vbDouble, _
             VbVarType.vbInteger, _
             VbVarType.vbLong, _
             VbVarType.vbLongLong, _
             VbVarType.vbDouble, _
             VbVarType.vbDecimal, _
             VbVarType.vbByte, _
             VbVarType.vbCurrency
            JSONString = CStr(Data)

        Case VbVarType.vbNull
            JSONString = "null"

        Case VbVarType.vbObject
            Dim Indentation As String: Indentation = String(Level * Indent, " ")
            If Utils.IsDictionary(Data) Then
                Dim StartObject As String: StartObject = "{" & vbNewLine & Indentation
                JSONString = JSONString & StartObject
                Dim Key As Variant
                For Each Key In Data.Keys
                    If JSONString <> StartObject Then
                        JSONString = JSONString & ", " & vbNewLine & Indentation
                    End If
                    JSONString = JSONString & """" & Key & """" & ": " & SerializeValue(Data(Key), Indent, Level + 1)
                Next
                Dim EndObject As String
                EndObject = vbNewLine & Strings.Left(Indentation, Level * Indent - Indent) & "}"
                JSONString = JSONString & EndObject
            ElseIf Utils.IsCollection(Data) Then
                Dim StartArray As String: StartArray = "[" & vbNewLine & Indentation
                JSONString = JSONString & StartArray
                Dim i As Long
                For i = 1 To Data.Count
                    If JSONString <> StartArray Then
                        JSONString = JSONString & ", " & vbNewLine & Indentation
                    End If
                    JSONString = JSONString & SerializeValue(Data(i), Indent, Level + 1)
                Next
                Dim EndArray As String
                EndArray = vbNewLine & Strings.Left(Indentation, Level * Indent - Indent) & "]"
                JSONString = JSONString & EndArray
            End If
    End Select
    SerializeValue = JSONString
End Function

Private Property Get Current() As String
    Current = Strings.Mid(this.JSON, this.Position, 1)
End Property

Private Sub NextChar()
    this.Position = this.Position + 1
    SkipWhitespaces
End Sub

Private Function ParseValue() As Variant
    If this.Position >= Strings.Len(this.JSON) Then Exit Function

    Select Case Current
        Case "{"
            Set ParseValue = ParseObject()

        Case "["
            Set ParseValue = ParseArray()

        Case "0" To "9"
            ParseValue = ParseNumber()

        Case """"
            ParseValue = ParseString()

        Case "t", "f", "n" ' true, false, null
            ParseValue = ParseLiteral()

        Case Else
            Information.Err.Raise _
                5, _
                "PJSON", _
                PStrings.FString("Invalid character '{0}' in position {1}", Current, this.Position)

    End Select
End Function

Private Sub SkipWhitespaces()
    If this.Position >= Strings.Len(this.JSON) Then Exit Sub
    Do While PStrings.IsWhiteSpace(Current)
        NextChar
        If this.Position >= Strings.Len(this.JSON) Then Exit Sub
    Loop
End Sub

#If DEV Then
  Private Function ParseObject() As Dictionary
#Else
  Private Function ParseObject() As Object
#End If
    NextChar
  #If DEV Then
    Dim Container As Dictionary: Set Container = NewDictionary()
  #Else
    Dim Container As Object: Set Container = NewDictionary()
  #End If
    Do While Current <> "}"
        If Current = "," Then
            NextChar
        End If

        Dim Key As String: Key = ParseValue()
        If Current = ":" Then
            NextChar
            Container.Add Key, ParseValue()
        End If
    Loop
    NextChar
    Set ParseObject = Container
End Function

Private Function ParseArray() As Collection
    Dim Container As Collection: Set Container = New Collection
    NextChar
    Do While Current <> "]"
        If Current = "," Then
            NextChar
        End If

        Container.Add ParseValue()
    Loop
    NextChar
    Set ParseArray = Container
End Function

Private Function ParseNumber() As Variant
    Do While Information.IsNumeric(Current) Or _
             Current = "."
        Dim Value As String: Value = Value & Current
        NextChar
    Loop
    ParseNumber = Conversion.Val(Value)
End Function

Private Function ParseString() As String
    NextChar
    Dim IsEscape As Boolean
    Do While Current <> """"
        IsEscape = Current = "\"
        If IsEscape Then
            this.Position = this.Position + 1
        End If
        Dim Value As String: Value = Value & Current
        this.Position = this.Position + 1
    Loop
    NextChar
    ParseString = Value
End Function

Private Function ParseLiteral() As Variant
    Do While Current Like "[tTrRuUeEfFaAlLsSnNuU]"
        Dim Value As String: Value = Value & Strings.LCase(Current)
        NextChar
    Loop
    If Value = "null" Then
        ParseLiteral = Null
    ElseIf Value = "true" Then
        ParseLiteral = True
    ElseIf Value = "false" Then
        ParseLiteral = False
    Else
        Information.Err.Raise 5, "PJSON", "Unexpected token " & Value
    End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PearPMProject.src.Config.ConfigIO"
Option Explicit

Private Type TConfigReader
    ConfigPath As String
End Type

Private this As TConfigReader

Public Property Get ConfigPath() As String
    ConfigPath = this.ConfigPath
End Property
Public Property Let ConfigPath(ByVal RHS As String)
    this.ConfigPath = RHS
End Property

Public Property Get Text() As String
    Text = Strings.Join(ReadLines(), vbNewLine)
End Property

#If DEV Then
  Public Function Read() As Dictionary
#Else
  Public Function Read() As Object
#End If
    Dim Lines As Variant: Lines = ReadLines()
    If Information.IsEmpty(Lines) Then Lines = Array()
  #If DEV Then
    Dim Data As Dictionary: Set Data = NewDictionary(VbCompareMethod.vbTextCompare)
  #Else
    Dim Data As Object: Set Data = NewDictionary(VbCompareMethod.vbTextCompare)
  #End If

    Dim Line As Variant
    For Each Line In Lines
        If PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), ";") Or _
           PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), "#") Or _
           PStrings.IsEqual(PStrings.FirstNonEmptyChar(Line), "") Then
            GoTo Continue
        End If
        Dim KeyValuePair As Variant
        KeyValuePair = Strings.Split(Line, "=")
        Data(Strings.Trim(KeyValuePair(0))) = Strings.Trim(KeyValuePair(1))
Continue:
    Next

    Set Read = Data
End Function

Public Sub CreateOrWrite(ByVal Key As String, ByVal Value As String)
  #If DEV Then
    Dim Data As Dictionary: Set Data = Me.Read()
  #Else
    Dim Data As Object: Set Data = Me.Read()
  #End If
    Data(Key) = Value
    WriteLines Data
End Sub

Public Sub DeleteKey(ByVal Key As String)
  #If DEV Then
    Dim Data As Dictionary: Set Data = Me.Read()
  #Else
    Dim Data As Object: Set Data = Me.Read()
  #End If
    If Data.Exists(Key) Then
        Data.Remove Key
        WriteLines Data
    End If
End Sub

#If DEV Then
  Public Sub AddDataWithoutReplace(ByRef Data As Dictionary)
#Else
  Public Sub AddDataWithoutReplace(ByRef Data As Object)
#End If
  #If DEV Then
    Dim Current As Dictionary: Set Current = Me.Read()
  #Else
    Dim Current As Object: Set Current = Me.Read()
  #End If

    Dim Key As Variant
    For Each Key In Data
        If Not Current.Exists(Key) Then
            Dim Value As Variant: Value = Data(Key).Default
            If Information.IsMissing(Value) Then GoTo Continue
            If Information.IsNull(Value) Then GoTo Continue
            If PStrings.StartsWith(Data(Key).Key, "_") Then GoTo Continue
            If Data(Key).DataType = VbVarType.vbBoolean Then GoTo Continue
            If Strings.Len(Value) = 0 Then GoTo Continue
            Current.Item(Key) = Value
        End If
Continue:
    Next
    WriteLines Current
End Sub

Private Function ReadLines() As Variant
    If Strings.Len(this.ConfigPath) = 0 Then Information.Err.Raise 76, Information.TypeName(Me)

    With NewFileSystemObject()
        If Not .FileExists(this.ConfigPath) Then
            .CreateTextFile this.ConfigPath
        End If
    End With

    With NewStream()
        .Charset = "UTF-8"
        .Open
        .LoadFromFile this.ConfigPath
        ReadLines = Strings.Split(.ReadText(), vbNewLine)
    End With
End Function

#If DEV Then
  Private Sub WriteLines(ByRef Data As Dictionary)
#Else
  Private Sub WriteLines(ByRef Data As Object)
#End If
    If Strings.Len(this.ConfigPath) = 0 Then Information.Err.Raise 76, Information.TypeName(Me)
    Dim Content As String
    Dim Key As Variant
    For Each Key In Data
        Dim Pair As String: Pair = PStrings.FString("{0}={1}", Key, Data(Key))
        If Strings.Len(Content) = 0 Then
            Content = Pair
        Else
            Content = PStrings.FString("{0}\\n{1}", Content, Pair)
        End If
    Next
    SaveToFile this.ConfigPath, Content
End Sub

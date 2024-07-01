Attribute VB_Name = "Locale"
'@Folder "PearPMProject.src.Utils"
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
      ByVal Locale As Long, _
      ByVal LCType As Long, _
      ByVal lpLCData As String, _
      ByVal cchData As Long) As Long
#Else
  Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
      ByVal Locale As Long, _
      ByVal LCType As Long, _
      ByVal lpLCData As String, _
      ByVal cchData As Long) As Long
#End If
Private Const LOCALE_IDEFAULTANSICODEPAGE As Long = &H1004

Public Function GetAnsiCodePage() As String
    Dim Buffer As String
    Buffer = Strings.Space(6)
    
    Dim LangID As Long: LangID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

    Dim Ret As Long
    Ret = GetLocaleInfo(LangID, LOCALE_IDEFAULTANSICODEPAGE, Buffer, Strings.Len(Buffer))
    If Ret > 0 Then
        GetAnsiCodePage = Strings.Left(Buffer, Ret - 1) ' Exclude null character
    Else
        GetAnsiCodePage = ""
    End If
End Function

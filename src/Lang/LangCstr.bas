Attribute VB_Name = "LangCstr"
'@Folder "PearPMProject.src.Lang"
Option Explicit

Public Function NewLang(ByVal Language As Languages) As Lang
    Set NewLang = New Lang
    NewLang.SetLang Language
End Function

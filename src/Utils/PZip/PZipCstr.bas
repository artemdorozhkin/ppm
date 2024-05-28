Attribute VB_Name = "PZipCstr"
'@Folder "PearPMProject.src.Utils.PZip"
Option Explicit

Public Function NewPZip(ByVal ZipFilePath As String)
    Set NewPZip = New PZip
    NewPZip.SetZip ZipFilePath
End Function

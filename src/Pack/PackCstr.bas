Attribute VB_Name = "PackCstr"
'@Folder "PearPMProject.src.Pack"
Option Explicit

Public Function NewPack(ByRef Project As Project) As Pack
    Set NewPack = New Pack
    Set NewPack.Project = Project
End Function

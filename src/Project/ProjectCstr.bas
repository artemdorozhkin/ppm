Attribute VB_Name = "ProjectCstr"
'@Folder "PearPMProject.src.Project"
Option Explicit

Public Function NewProject(ByVal Project As VBProject) As Project
    Set NewProject = New Project
    Set NewProject.Project = Project
End Function

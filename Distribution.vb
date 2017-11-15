Imports System.Text

Public Class Distribution
    Dim ini As IniFile

    Public Sub DistributeProject(ByVal ProjectPath As String)
        ini = New IniFile()
        ini.Load(ProjectPath)

    End Sub
End Class

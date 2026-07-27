Imports System
Imports System.IO

Namespace Sage300

    ''' Keeps a copy of the sample Excel file for each column-mapping template in a "Plantillas"
    ''' folder next to the application executable, so the column links can be shown again after
    ''' restarting the app without needing the original file to still exist in its original location.
    Public Class TemplateSampleStore

        Private Shared Function GetFolder() As String
            Dim folder = Path.Combine(AppContext.BaseDirectory, "Plantillas")
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function GetSamplePath(templateName As String) As String
            Return Path.Combine(GetFolder(), templateName & ".xlsx")
        End Function

        Public Shared Function TryGetExistingSamplePath(templateName As String) As String
            Dim filePath = GetSamplePath(templateName)
            Return If(File.Exists(filePath), filePath, Nothing)
        End Function

        Public Shared Sub SaveSample(templateName As String, sourceFilePath As String)
            File.Copy(sourceFilePath, GetSamplePath(templateName), overwrite:=True)
        End Sub

        Public Shared Sub DeleteSample(templateName As String)
            Dim filePath = GetSamplePath(templateName)
            If File.Exists(filePath) Then File.Delete(filePath)
        End Sub

    End Class

End Namespace

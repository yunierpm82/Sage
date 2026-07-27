Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.Json

Namespace Sage300

    ''' Guarda/lee plantillas de mapeo de columnas (columna requerida -> columna real del Excel
    ''' de una aplicacion de terceros), una por archivo JSON, en la carpeta de datos del usuario.
    Public Class ColumnMappingStore

        Private Shared Function GetFolder() As String
            Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage", "ColumnMappings")
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function ListProfiles() As List(Of String)
            Return Directory.GetFiles(GetFolder(), "*.json").
                Select(Function(f) Path.GetFileNameWithoutExtension(f)).
                OrderBy(Function(n) n, StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Public Shared Function Load(profileName As String) As Dictionary(Of String, String)
            Dim filePath = Path.Combine(GetFolder(), profileName & ".json")
            If Not File.Exists(filePath) Then Return New Dictionary(Of String, String)
            Dim json = File.ReadAllText(filePath)
            Return JsonSerializer.Deserialize(Of Dictionary(Of String, String))(json)
        End Function

        Public Shared Sub Save(profileName As String, mapping As Dictionary(Of String, String))
            Dim filePath = Path.Combine(GetFolder(), profileName & ".json")
            Dim json = JsonSerializer.Serialize(mapping, New JsonSerializerOptions With {.WriteIndented = True})
            File.WriteAllText(filePath, json)
        End Sub

        Public Shared Sub Delete(profileName As String)
            Dim filePath = Path.Combine(GetFolder(), profileName & ".json")
            If File.Exists(filePath) Then File.Delete(filePath)
        End Sub

    End Class

End Namespace

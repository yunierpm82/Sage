Imports System
Imports System.Collections.Generic
Imports ClosedXML.Excel

Namespace Sage300

    ''' Formato esperado del Excel (primera fila = encabezados, sin importar mayusculas/minusculas):
    ''' Entrada | Fecha | Cuenta | Descripcion | Referencia | Debito | Credito
    ''' Las filas con el mismo numero de "Entrada" forman un solo asiento (journal entry)
    ''' con varias lineas dentro del mismo lote (batch).
    Public Class TransactionLine
        Public Property EntryNumber As Integer
        Public Property EntryDate As Date
        Public Property Account As String
        Public Property Description As String
        Public Property Reference As String
        Public Property Debit As Decimal
        Public Property Credit As Decimal
    End Class

    Public Class ExcelTransactionReader

        Private Shared ReadOnly RequiredColumns As String() = {
            "ENTRADA", "FECHA", "CUENTA", "DESCRIPCION", "REFERENCIA", "DEBITO", "CREDITO"
        }

        Public Shared Function ReadEntries(filePath As String) As List(Of TransactionLine)
            Dim result As New List(Of TransactionLine)

            Using workbook As New XLWorkbook(filePath)
                Dim worksheet = workbook.Worksheet(1)
                Dim headerRow = worksheet.FirstRowUsed()
                If headerRow Is Nothing Then
                    Throw New Exception("El archivo Excel está vacío.")
                End If

                Dim columnIndex As New Dictionary(Of String, Integer)
                For Each cell In headerRow.CellsUsed()
                    columnIndex(cell.GetString().Trim().ToUpperInvariant()) = cell.Address.ColumnNumber
                Next

                For Each columnName In RequiredColumns
                    If Not columnIndex.ContainsKey(columnName) Then
                        Throw New Exception($"El archivo Excel debe tener una columna '{columnName}'. Columnas encontradas: {String.Join(", ", columnIndex.Keys)}")
                    End If
                Next

                Dim lastRowUsed = worksheet.LastRowUsed()
                If lastRowUsed Is Nothing Then
                    Return result
                End If

                For rowNumber = headerRow.RowNumber() + 1 To lastRowUsed.RowNumber()
                    Dim row = worksheet.Row(rowNumber)
                    If row.IsEmpty() Then Continue For

                    Dim line As New TransactionLine With {
                        .EntryNumber = CInt(GetNumberOrZero(row.Cell(columnIndex("ENTRADA")))),
                        .EntryDate = GetDateOrToday(row.Cell(columnIndex("FECHA"))),
                        .Account = row.Cell(columnIndex("CUENTA")).GetString().Trim(),
                        .Description = row.Cell(columnIndex("DESCRIPCION")).GetString().Trim(),
                        .Reference = row.Cell(columnIndex("REFERENCIA")).GetString().Trim(),
                        .Debit = GetNumberOrZero(row.Cell(columnIndex("DEBITO"))),
                        .Credit = GetNumberOrZero(row.Cell(columnIndex("CREDITO")))
                    }

                    If String.IsNullOrWhiteSpace(line.Account) Then Continue For

                    result.Add(line)
                Next
            End Using

            Return result
        End Function

        Private Shared Function GetNumberOrZero(cell As IXLCell) As Decimal
            If cell.IsEmpty() Then Return 0D
            Return CDec(cell.GetDouble())
        End Function

        Private Shared Function GetDateOrToday(cell As IXLCell) As Date
            If cell.IsEmpty() Then Return Date.Today
            Return cell.GetDateTime()
        End Function

    End Class

End Namespace

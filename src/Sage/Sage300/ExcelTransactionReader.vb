Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports ClosedXML.Excel

Namespace Sage300

    ''' Expected Excel format (first row = headers, case-insensitive):
    ''' Entry | Date | Account | Description | Reference | TransAmt
    ''' TransAmt: positive = debit, negative = credit (passed straight to SCURNAMT with its sign).
    ''' Rows sharing the same "Entry" number form a single journal entry with multiple lines
    ''' inside the same batch.
    '''
    ''' If the real Excel (from a third-party application) uses different column names, a
    ''' "columnMapping" (required column -> real column) generated from the GL Booking >
    ''' Relate Columns screen can be passed in.
    Public Class TransactionLine
        Public Property EntryNumber As Integer
        Public Property EntryDate As Date
        Public Property Account As String
        Public Property Description As String
        Public Property Reference As String
        Public Property Amount As Decimal
    End Class

    Public Class ExcelTransactionReader

        Public Shared ReadOnly RequiredColumnKeys As String() = {
            "Entry", "Date", "Account", "Description", "Reference", "TransAmt"
        }

        Public Shared Function ReadHeaders(filePath As String) As List(Of String)
            Using workbook As New XLWorkbook(filePath)
                Dim worksheet = workbook.Worksheet(1)
                Dim headerRow = worksheet.FirstRowUsed()
                If headerRow Is Nothing Then Return New List(Of String)

                Return headerRow.CellsUsed().
                    Select(Function(c) c.GetString().Trim()).
                    Where(Function(s) s <> "").
                    ToList()
            End Using
        End Function

        Public Shared Function ReadEntries(filePath As String, Optional columnMapping As Dictionary(Of String, String) = Nothing) As List(Of TransactionLine)
            Dim result As New List(Of TransactionLine)

            Using workbook As New XLWorkbook(filePath)
                Dim worksheet = workbook.Worksheet(1)
                Dim headerRow = worksheet.FirstRowUsed()
                If headerRow Is Nothing Then
                    Throw New Exception("The Excel file is empty.")
                End If

                Dim columnIndex As New Dictionary(Of String, Integer)
                For Each cell In headerRow.CellsUsed()
                    columnIndex(cell.GetString().Trim().ToUpperInvariant()) = cell.Address.ColumnNumber
                Next

                Dim resolvedColumn As New Dictionary(Of String, Integer)
                For Each requiredKey In RequiredColumnKeys
                    Dim targetHeader = requiredKey

                    If columnMapping IsNot Nothing Then
                        Dim mappedName = columnMapping.
                            Where(Function(kv) String.Equals(kv.Key, requiredKey, StringComparison.OrdinalIgnoreCase)).
                            Select(Function(kv) kv.Value).
                            FirstOrDefault()
                        If Not String.IsNullOrWhiteSpace(mappedName) Then targetHeader = mappedName
                    End If

                    Dim lookupKey = targetHeader.Trim().ToUpperInvariant()
                    If Not columnIndex.ContainsKey(lookupKey) Then
                        Throw New Exception($"Could not find the column '{targetHeader}' (needed for '{requiredKey}') in the Excel file. Columns found: {String.Join(", ", columnIndex.Keys)}. Use GL Booking > Relate Columns to map the correct names.")
                    End If

                    resolvedColumn(requiredKey.ToUpperInvariant()) = columnIndex(lookupKey)
                Next

                Dim lastRowUsed = worksheet.LastRowUsed()
                If lastRowUsed Is Nothing Then
                    Return result
                End If

                For rowNumber = headerRow.RowNumber() + 1 To lastRowUsed.RowNumber()
                    Dim row = worksheet.Row(rowNumber)
                    If row.IsEmpty() Then Continue For

                    Dim line As New TransactionLine With {
                        .EntryNumber = CInt(GetNumberOrZero(row.Cell(resolvedColumn("ENTRY")))),
                        .EntryDate = GetDateOrToday(row.Cell(resolvedColumn("DATE"))),
                        .Account = row.Cell(resolvedColumn("ACCOUNT")).GetString().Trim(),
                        .Description = row.Cell(resolvedColumn("DESCRIPTION")).GetString().Trim(),
                        .Reference = row.Cell(resolvedColumn("REFERENCE")).GetString().Trim(),
                        .Amount = GetNumberOrZero(row.Cell(resolvedColumn("TRANSAMT")))
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

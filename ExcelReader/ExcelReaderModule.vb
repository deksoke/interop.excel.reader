Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Excel

Public Module ExcelReaderModule

    <Extension()> _
    Public Function ChangeDataType(ByVal dt As System.Data.DataTable, ByVal colName As String, ByVal targetDataType As Type, ByRef dtReturn As System.Data.DataTable) As Boolean
        dtReturn = dt.Clone
        Try
            dtReturn.Columns(colName).DataType = targetDataType
            For Each dr As System.Data.DataRow In dt.Rows
                dtReturn.ImportRow(dr)
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    <Extension()> _
    Public Function ToDataSet(ByVal xlsWorkBook As Workbook, Optional ByVal CloneColumnHeader As Boolean = False, Optional ByVal AutoChangeColumnDataType As Boolean = False) As System.Data.DataSet
        Dim ds As New System.Data.DataSet

        ' Loop over all sheets.
        For i As Integer = 1 To xlsWorkBook.Sheets.Count
            ' Get sheet.
            Dim sheet As Worksheet = xlsWorkBook.Sheets(i)
            Dim sheetName As String = sheet.Name 'tableName
            ds.Tables.Add(sheet.ToDataTable(CloneColumnHeader, AutoChangeColumnDataType))
        Next

        Return ds
    End Function

    <Extension()> _
    Public Function ToDataTable(ByVal xlsWorkSheet As Worksheet, Optional ByVal CloneColumnHeader As Boolean = False, Optional ByVal AutoChangeColumnDataType As Boolean = False) As System.Data.DataTable
        ' Get sheet.
        Dim sheetName As String = xlsWorkSheet.Name 'tableName
        Dim dt As New System.Data.DataTable(sheetName)
        Dim dtClone As System.Data.DataTable = Nothing

        ' Get range.
        Dim r As Range = xlsWorkSheet.UsedRange

        Dim colNum, rowNum As Integer
        colNum = r.Columns.Count
        rowNum = r.Rows.Count

        ' Load all cells into 2d array.
        Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)
        If array IsNot Nothing Then

            Dim FirstDataRowArray(colNum - 1) As Object
            Dim IsAssignFirstDataRow As Boolean = False

            Dim HeaderFlag As Boolean = CloneColumnHeader
            For j As Integer = 0 To rowNum - 1
                Dim RawDataRowArray(colNum - 1) As Object
                Dim cellValue As Object
                For k As Integer = 0 To colNum - 1
                    cellValue = array(j + 1, k + 1)
                    If (cellValue Is Nothing) Then
                        cellValue = ""
                    End If
                    RawDataRowArray(k) = cellValue
                Next

                If HeaderFlag Then
                    AddValuesToTable(RawDataRowArray, dt, HeaderFlag)
                    HeaderFlag = False
                Else
                    If IsAssignFirstDataRow = False Then
                        IsAssignFirstDataRow = True
                        FirstDataRowArray = RawDataRowArray
                    End If
                    AddValuesToTable(RawDataRowArray, dt)
                End If
            Next

            If AutoChangeColumnDataType And IsAssignFirstDataRow Then
                dtClone = dt.Clone
                Dim dcName As String
                Dim t1, t2 As Type
                For i As Integer = 0 To dt.Columns.Count - 1
                    t1 = dt.Columns(i).DataType
                    t2 = FirstDataRowArray(i).GetType
                    If Not t1.Equals(t2) Then
                        dcName = dt.Columns(i).ColumnName
                        dt.ChangeDataType(dcName, t2, dtClone)
                        dt = dtClone.Copy
                    End If
                Next
            End If
        End If

        Return dt
    End Function

    Private Function AddValuesToTable(ByRef source() As Object, ByVal destination As System.Data.DataTable, Optional ByVal HeaderFlag As Boolean = False) As Boolean
        'Ensures a datatable can hold an array of values and then adds a new row 
        Try
            Dim existing As Integer = destination.Columns.Count
            If HeaderFlag Then
                Resolve_Duplicate_Names(source)
                For i As Integer = 0 To source.Length - existing - 1
                    If source(i).ToString.Trim = "" Then source(i) = ""
                    destination.Columns.Add(source(i).ToString, Type.GetType("System.String")) 'GetType(String))
                Next i
                Return True
            End If
            For i As Integer = 0 To source.Length - existing - 1
                destination.Columns.Add("Column" & (existing + 1 + i).ToString, Type.GetType("System.String"))  'GetType(String))
            Next
            destination.Rows.Add(source)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub Resolve_Duplicate_Names(ByRef source() As Object)
        ' Resolves the possibility of duplicated names by appending "Duplicate Name" and a number at the end of any duplicates
        Dim i, n, dnum As Integer
        dnum = 1
        For n = 0 To source.Length - 1
            For i = n + 1 To source.Length - 1
                If source(i) = source(n) Then
                    source(i) = source(i) & "Duplicate Name " & dnum
                    dnum += 1
                End If
            Next
        Next
        Return
    End Sub

    Public Class ExcelReader
        Implements IDisposable

        Private _xlsApp As Application = Nothing
        Private _xlsWorkBook As Workbook = Nothing

        Public Sub New(ByVal FilePath As String)
            If System.IO.File.Exists(FilePath) Then
                _xlsApp = New Application
                _xlsWorkBook = _xlsApp.Workbooks.Open(FilePath)
            End If
        End Sub

        Public Function GetWorkBook() As Workbook
            Return _xlsWorkBook
        End Function

        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free other state (managed objects).
                End If

                ' TODO: free your own state (unmanaged objects).
                ' TODO: set large fields to null.
                If _xlsWorkBook IsNot Nothing Then
                    _xlsWorkBook.Close(False)
                    _xlsWorkBook = Nothing
                End If
                If _xlsApp IsNot Nothing Then
                    _xlsApp.Workbooks.Close()
                    _xlsApp.Quit()
                    _xlsApp = Nothing
                End If
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Module

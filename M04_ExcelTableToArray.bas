Attribute VB_Name = "M04_ExcelTableToArray"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date: 19/09/2024
' Description: Obtain Excel Table object as Array.
' =======================================================
Option Explicit

Function fExcelTableToList() As Variant
    Dim sXlsPath        As String               ' Ruta del fichero
    Dim oXlsApp         As Excel.Application    ' Instancia de Aplicacion de Excel
    Dim oXlsWorkbook    As Excel.Workbook       ' Instancia de Workbook de Excel
    Dim oXlsTable       As Variant              ' Array con el contenido de la tabla
    
    Let sXlsPath = "C:\Users\pablo\Desktop\testSources.xlsx"
    Set oXlsApp = CreateObject("Excel.Application")
    Set oXlsWorkbook = oXlsApp.Workbooks.Open(FileName:=sXlsPath, ReadOnly:=True)
    With oXlsApp
        With oXlsWorkbook
            .Activate
            oXlsTable = fGetTableAsArray(.Sheets("Sheet1"), "Ficheros")
        End With
        .Quit
    End With
    fExcelTableToList = oXlsTable
End Function

Private Function fGetTableAsArray(Sheet As Excel.Worksheet, TableName As String) As Variant
    Dim tbl As ListObject
    Dim arr As Variant
    Dim i   As Long
    Dim j   As Long

    Set tbl = Sheet.ListObjects(TableName)
    If tbl Is Nothing Then
        MsgBox "Table '" & TableName & "' not found on sheet '" & Sheet.Name & "'."
        Exit Function
    End If
    ReDim arr(1 To tbl.ListRows.Count, 1 To tbl.ListColumns.Count)
    For i = 1 To tbl.ListRows.Count
        For j = 1 To tbl.ListColumns.Count
            arr(i, j) = tbl.ListRows(i).Range.Cells(1, j).Value
        Next j
    Next i
    fGetTableAsArray = arr
End Function

Attribute VB_Name = "M04_ExcelTableToArray"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date: 19/09/2024
' Description: Obtain Excel Table object as Array.
' =======================================================
Option Explicit

Function fExcelTableToList() As Variant
    Dim oXlsApp         As Excel.Application    ' Instancia de Aplicacion de Excel
    Dim oXlsWorkbook    As Excel.Workbook       ' Instancia de Workbook de Excel
    Dim vXlsTable       As Variant              ' Array con el contenido de la tabla
    
    Set oXlsApp = CreateObject("Excel.Application")
    Set oXlsWorkbook = oXlsApp.Workbooks.Open(FileName:=gsEXCEL_PATH, ReadOnly:=True)
    With oXlsApp
        With oXlsWorkbook
            .Activate
            vXlsTable = fGetTableAsArray(.Sheets(gsEXCEL_SHEET_NAME), gsEXCEL_TABLE_NAME)
        End With
        .Quit
    End With
    fExcelTableToList = vXlsTable
End Function

Private Function fGetTableAsArray(Sheet As Excel.Worksheet, TableName As String) As Variant
    Dim loTbl   As ListObject
    Dim arr     As Variant
    Dim i       As Long
    Dim j       As Long

    Set loTbl = Sheet.ListObjects(TableName)
    If loTbl Is Nothing Then
        MsgBox "Table '" & TableName & "' not found on sheet '" & Sheet.Name & "'."
        Exit Function
    End If
    ReDim arr(1 To loTbl.ListRows.Count, 1 To loTbl.ListColumns.Count)
    For i = 1 To loTbl.ListRows.Count
        For j = 1 To loTbl.ListColumns.Count
            arr(i, j) = loTbl.ListRows(i).Range.Cells(1, j).Value
        Next j
    Next i
    fGetTableAsArray = arr
End Function

Attribute VB_Name = "M01_MailsToWord"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date:
' Description: Programa para copiar el texto de interes
'   de un email en un word. Modulo principal (main)
' =======================================================
Option Explicit
Public gsSelectedFile    As String
Public gbIsCancelled      As Boolean
'Ruta donde se encuentra el Excel con la tabla de archivos de destino
Public Const gsEXCEL_PATH  As String = '<RUTA_ABSOLUTA_AL_EXCEL>
'Nombre de la hoja de Excel con la tabla de archivos de destino
Public Const gsEXCEL_SHEET_NAME As String = '<NOMBRE_HOJA_EXCEL>
'Nombre del objeto tabla que contiene las rutas de destino
Public Const gsEXCEL_TABLE_NAME As String = '<NOMBRE_OBJETO_TABLA_EXCEL>

Sub sMailToWord()
    Dim dContent    As New Scripting.Dictionary
    
    Set dContent = CreateObject("Scripting.Dictionary")
    UserForm1.Show
    If gbIsCancelled Then
        Exit Sub
    ElseIf gsSelectedFile = "" Then
        MsgBox _
            Title:="ERROR - Fichero no seleccionado", _
            Prompt:="Seleccione un fichero de destino para que el proceso se ejecute correctamente"
        Exit Sub
    End If
    sMailDict dContent
    sAppendToWord dContent, gsSelectedFile
End Sub

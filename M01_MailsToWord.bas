Attribute VB_Name = "M01_MailsToWord"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date:
' Description: Programa para copiar el texto de interes
'   de un email en un word.
' =======================================================
Option Explicit
Public sSelectedFile    As String
Public isCancelled      As Boolean

Sub sMailToWord()
    Dim dContent    As New Scripting.Dictionary
    Dim Path        As String
    
    Set dContent = CreateObject("Scripting.Dictionary")
    UserForm1.Show
    If isCancelled Then
        Exit Sub
    End If
    sMailDict dContent
    sAppendToWord dContent, sSelectedFile
End Sub

Attribute VB_Name = "M03_AppendToWord"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date: 19/09/2024
' Description: Añade elementos seleccionados a un word.
' =======================================================
Option Explicit
Option Private Module

Public Sub sAppendToWord(ByRef dContent As Scripting.Dictionary, _
    ByVal DocPath As String _
)
    Dim i                       As Integer
    Dim wdApp                   As word.Application
    Dim wdDoc                   As word.document
    Dim aArrOfStrs(5)           As String
    
    Let i = 0
    sOpenWord wdApp, wdDoc, DocPath
    fFillArr aArrOfStrs, dContent
    While i <= UBound(aArrOfStrs)
        sInsertTextAtEndOfDocument aArrOfStrs(i), wdDoc, wdApp
        i = i + 1
    Wend
    sSaveAndClose wdApp, wdDoc, DocPath
End Sub


Private Sub sOpenWord( _
    ByRef wdApp As word.Application, _
    ByRef wdDoc As word.document, _
    ByVal wdPath As String _
)
    ' Description: Open or create a Word file
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(wdPath, Visible:=True)
    wdDoc.Activate
End Sub

Private Sub fFillArr( _
    ByRef aArrOfStrs() As String, _
    ByVal dContent As Scripting.Dictionary _
)
    aArrOfStrs(0) = "-----------------------"
    aArrOfStrs(1) = "From: " & dContent("SenderName") & _
        "(" & dContent("SenderEmailAddress") & ")"
    aArrOfStrs(2) = "Sent: " & dContent("Sent")
    aArrOfStrs(3) = "To: " & dContent("To")
    aArrOfStrs(4) = "Subject: " & dContent("Subject")
    aArrOfStrs(5) = dContent("SelectedText")
End Sub

'Append text at the end of the document
Private Sub sInsertTextAtEndOfDocument( _
    ByVal textToAppend As String, _
    ByVal wdDoc As word.document, _
    ByVal wdApp As word.Application _
)
    Dim isAtBeginningOfLine As Boolean
    Dim isAtEndOfLine       As Boolean
    
    wdApp.Selection.EndKey Unit:=wdStory
    Let isAtBeginningOfLine = (wdApp.Selection.Start = wdApp.Selection.Paragraphs(1).Range.Start)
    Let isAtEndOfLine = (wdApp.Selection.End = wdApp.Selection.Paragraphs(1).Range.End - 1)
    If isAtEndOfLine And Not isAtBeginningOfLine Then
        wdDoc.content.InsertAfter text:=vbCrLf
    End If
    wdDoc.content.InsertAfter text:=textToAppend & vbCrLf
End Sub

' Save changes done in `wdDoc` as `wdPath` and close Word app and document
Private Sub sSaveAndClose( _
    ByRef wdApp As word.Application, _
    ByRef wdDoc As word.document, _
    ByVal wdPath As String _
)
    wdDoc.SaveAs2 FileName:=wdPath
    wdApp.Selection.EndKey Unit:=wdStory
    wdApp.Visible = True
    wdApp.Activate
    'wdDoc.Close
    'wdApp.Quit
End Sub



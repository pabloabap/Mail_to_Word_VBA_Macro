Attribute VB_Name = "M03_AppendToWord"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date: 19/09/2024
' Description: A�ade elementos seleccionados a un word.
' =======================================================
Option Explicit
Option Private Module

Public Sub sAppendToWord(ByRef dContent As Scripting.Dictionary, _
    ByVal DocPath As String _
)
    Dim i                       As Integer
    Dim oWdApp                  As word.Application
    Dim oWdDoc                  As word.document
    Dim aArrOfStrs(5)           As String
    
    Let i = 0
    sOpenWord oWdApp, oWdDoc, DocPath
    fFillArr aArrOfStrs, dContent
    While i <= UBound(aArrOfStrs)
        sInsertTextAtEndOfDocument aArrOfStrs(i), oWdDoc, oWdApp
        i = i + 1
    Wend
    sSaveAndClose oWdApp, oWdDoc, DocPath
End Sub

Private Sub sOpenWord( _
    ByRef oWdApp As word.Application, _
    ByRef oWdDoc As word.document, _
    ByVal sWdPath As String _
)
    On Error Resume Next
        Set oWdApp = GetObject(, "Word.Application")
    On Error GoTo NO_WORD_FOUND
    If oWdApp Is Nothing Then
        GoTo NO_WORD_FOUND
    Else
        For Each oWdDoc In oWdApp.Documents
            If oWdDoc.FullName = sWdPath Then
                oWdApp.Visible = True
                oWdDoc.Activate
                Exit Sub
            End If
        Next
    End If
        GoTo NO_WORD_FOUND
    Exit Sub
NO_WORD_FOUND:
    If oWdApp Is Nothing Then
        Set oWdApp = CreateObject("Word.Application")
    End If
    Set oWdDoc = oWdApp.Documents.Open(sWdPath, Visible:=True)
    oWdApp.Visible = True
    oWdDoc.Activate
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
    ByVal sTextToAppend As String, _
    ByVal oWdDoc As word.document, _
    ByVal oWdApp As word.Application _
)
    Dim bIsAtBeginningOfLine As Boolean
    Dim bIsAtEndOfLine       As Boolean
    
    oWdApp.Selection.EndKey Unit:=wdStory
    Let bIsAtBeginningOfLine = (oWdApp.Selection.Start = oWdApp.Selection.Paragraphs(1).Range.Start)
    Let bIsAtEndOfLine = (oWdApp.Selection.End = oWdApp.Selection.Paragraphs(1).Range.End - 1)
    If bIsAtEndOfLine And Not bIsAtBeginningOfLine Then
        oWdDoc.content.InsertAfter text:=vbCrLf
    End If
    oWdDoc.content.InsertAfter text:=sTextToAppend & vbCrLf
End Sub

' Save changes done in `oWdDoc` as `sWdPath` and close Word app and document
Private Sub sSaveAndClose( _
    ByRef oWdApp As word.Application, _
    ByRef oWdDoc As word.document, _
    ByVal sWdPath As String _
)
    oWdDoc.SaveAs2 FileName:=sWdPath
    oWdApp.Selection.EndKey Unit:=wdStory
    oWdApp.Visible = True
    oWdApp.Activate
    ' Si queres cerrar automaticamente el documento con los cambios descomentar las dos lineas siguientes.
    'oWdDoc.Close
    'oWdApp.Quit
End Sub



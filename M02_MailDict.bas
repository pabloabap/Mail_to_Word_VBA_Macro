Attribute VB_Name = "M02_MailDict"
' =======================================================
' Author: Pablo Abad Aparicio  (pablo.abad96@gmail.com)
' Creation date:
' Description: Crea un diccionario con los elementos de
'               interes de un email
' =======================================================
Option Explicit
Option Private Module


Public Sub sMailDict(ByRef dContent As Scripting.Dictionary)
    Dim OutApp  As Object
    Dim OutMail As Object
    Dim olInsp  As Object
    Dim wdDoc   As Object
    Dim i       As Integer
    
    Set OutApp = GetObject(, "Outlook.Application")
    Set OutMail = OutApp.ActiveExplorer.Selection.Item(1)
    With OutMail
        Set olInsp = .GetInspector
        Set wdDoc = olInsp.WordEditor
        sFillDictContent dContent, wdDoc, OutMail
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set olInsp = Nothing
    Set wdDoc = Nothing
End Sub

Private Sub sFillDictContent( _
    ByRef dContent As Scripting.Dictionary, _
    ByVal wdDoc As Object, _
    ByVal OutMail As Object _
    )
    With OutMail
        dContent.Add "SenderName", .SenderName
        dContent.Add "SenderEmailAddress", .SenderEmailAddress
        dContent.Add "Sent", .SentOn
        dContent.Add "To", .To
        dContent.Add "Subject", .Subject
        dContent.Add "SelectedText", wdDoc.Application.Selection.Range.text
    End With
End Sub

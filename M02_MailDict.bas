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
    Dim oOutApp     As Object
    Dim oOutMail    As Object
    Dim oInsp       As Object
    Dim oWdDoc      As Object
    Dim i           As Integer
    
    Set oOutApp = GetObject(, "Outlook.Application")
    Set oOutMail = oOutApp.ActiveExplorer.Selection.Item(1)
    With oOutMail
        Set oInsp = .GetInspector
        Set oWdDoc = oInsp.WordEditor
        sFillDictContent dContent, oWdDoc, oOutMail
    End With
    Set oOutMail = Nothing
    Set oOutApp = Nothing
    Set oInsp = Nothing
    Set oWdDoc = Nothing
End Sub

Private Sub sFillDictContent( _
    ByRef dContent As Scripting.Dictionary, _
    ByVal oWdDoc As Object, _
    ByVal oOutMail As Object _
    )
    With oOutMail
        dContent.Add "SenderName", .SenderName
        dContent.Add "SenderEmailAddress", .SenderEmailAddress
        dContent.Add "Sent", .SentOn
        dContent.Add "To", .To
        dContent.Add "Subject", .Subject
        dContent.Add "SelectedText", oWdDoc.Application.Selection.Range.text
    End With
End Sub

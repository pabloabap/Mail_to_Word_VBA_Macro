VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3430
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10010
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub UserForm_Initialize()
    gbIsCancelled = False
    With ListBox1
        .List() = fExcelTableToList()
        .ColumnWidths = CStr(.Width * 0.2) & ";" & CStr(.Width * 0.8)
    End With
End Sub

Public Sub ListBox1_Click()
    Dim LBItem As Long
    For LBItem = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(LBItem) = True Then
            sSelectedFile = ListBox1.List(LBItem, 1)
        End If
    Next
End Sub

Private Sub btnCancel_Click()
    gbIsCancelled = True
    Unload UserForm1
End Sub

Private Sub CommandButton1_Click()
    Unload UserForm1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        gbIsCancelled = True
        Unload UserForm1
    End If
End Sub





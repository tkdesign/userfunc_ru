VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SortDirectionDialog 
   Caption         =   "?????????? ???? ?????? ??????????? ??????????"
   ClientHeight    =   2288
   ClientLeft      =   117
   ClientTop       =   468
   ClientWidth     =   5161
   OleObjectBlob   =   "SortDirectionDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SortDirectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ButtonPressed As Integer

Private Sub Cancel_Button_Click()
    DialogResult = 0
    Unload Me
End Sub

Private Sub OK_Button_Click()
    DialogResult = 1
    Me.Hide
End Sub

Public Property Let DialogResult(ByVal ButtonCode As Integer)
    ButtonPressed = ButtonCode
End Property
 
Public Property Get DialogResult() As Integer
    DialogResult = ButtonPressed
End Property

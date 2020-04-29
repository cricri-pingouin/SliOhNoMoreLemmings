VERSION 5.00
Begin VB.Form frmLemm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   6
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   3975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   18
      Min             =   3
      TabIndex        =   0
      Top             =   360
      Value           =   3
      Width           =   3975
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number: 1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label lblPage 
      Caption         =   "Page: 3"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmLemm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Calculate
End Sub

Private Sub HScroll1_Change()
    lblPage.Caption = "Page: " & HScroll1.Value
    Call Calculate
End Sub

Private Sub HScroll2_Change()
    lblNumber.Caption = "Number: " & HScroll2.Value
    Call Calculate
End Sub

Private Sub Calculate()
    Dim x As Long
    Dim i As Integer
    
    x = 10199 - 3463 - 15649 '15649=3463*5, to shift page, 3463 to shift number
    For i = 3 To HScroll1.Value
        x = x + 15649
        If x > 32000 Then x = x - 32000
    Next
    For i = 1 To HScroll2.Value
        x = x + 3463
        If x > 32000 Then x = x - 32000
    Next
    Me.Caption = "Your code is: " & x
End Sub

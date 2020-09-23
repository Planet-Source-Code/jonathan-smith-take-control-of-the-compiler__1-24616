VERSION 5.00
Begin VB.Form frmInterface 
   Caption         =   "In-Line Assembly Example Interface"
   ClientHeight    =   1815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbIterations 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   1515
   End
   Begin VB.TextBox tbOperand 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   540
      Width           =   1515
   End
   Begin VB.CommandButton cbDivideBy2ByShifting 
      Caption         =   "Divide by 2 by shifting"
      Height          =   405
      Left            =   1980
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
   Begin VB.CommandButton cbDivideBy2Normally 
      Caption         =   "Divide by 2 normally"
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label lblTimes 
      Alignment       =   1  'Right Justify
      Caption         =   "Iterations:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label lblOperand 
      Alignment       =   1  'Right Justify
      Caption         =   "Operand:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DivideBy2(sDirective As String)
    Dim lngOperand As Long, lngIterations As Long, lngAnswer As Long
    Dim timerStart As Single, timerEnd As Single
    If IsNumeric(tbOperand.Text) Then
        lngOperand = CLng(tbOperand.Text)
    Else
        MsgBox "You must enter a number in the operand box."
        Exit Sub
    End If
    If IsNumeric(tbIterations.Text) Then
        lngIterations = CLng(tbIterations.Text)
    Else
        lngIterations = 1
    End If
    If sDirective = "normally" Then
        timerStart = Timer
        lngAnswer = DivideBy2Normally(lngOperand, lngIterations)
    Else
        timerStart = Timer
        lngAnswer = DivideBy2ByShifting(lngOperand, lngIterations)
    End If
    timerEnd = Timer
    MsgBox "Answer: " & lngAnswer & " To divide " & lngOperand & " by two " & lngIterations & " times " & sDirective & " took: " & timerEnd - timerStart & " seconds."
End Sub

Private Sub cbDivideBy2ByShifting_Click()
    DivideBy2 "by shifting"
End Sub

Private Sub cbDivideBy2Normally_Click()
    DivideBy2 "normally"
End Sub


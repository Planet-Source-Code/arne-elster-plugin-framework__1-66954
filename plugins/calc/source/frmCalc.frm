VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Sum up values"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   750
      TabIndex        =   4
      Top             =   675
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   675
      Width           =   1140
   End
   Begin VB.TextBox txtVal2 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "1"
      Top             =   225
      Width           =   1365
   End
   Begin VB.TextBox txtVal1 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Text            =   "1"
      Top             =   225
      Width           =   1365
   End
   Begin VB.Label lblSign 
      AutoSize        =   -1  'True
      Caption         =   "+"
      Height          =   195
      Left            =   1650
      TabIndex        =   1
      Top             =   225
      Width           =   90
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnCancel    As Boolean

Private Sub cmdCancel_Click()
    blnCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    blnCancel = False
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    cmdCancel_Click
End Sub

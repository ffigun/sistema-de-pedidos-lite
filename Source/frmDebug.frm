VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFetch 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblMain 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblMain.Caption = _
        "TaskHolder   = " & vbNewLine & _
        "---------------------" & vbNewLine & _
        "CurrentTask  = " & vbNewLine & _
        "CurrentReport = "
End Sub

Private Sub tmrFetch_Timer()
    lblMain.Caption = _
        "TaskHolder   = " & TaskHolder & vbNewLine & _
        "---------------------" & vbNewLine & _
        "CurrentTask  = " & CurrentTask & vbNewLine & _
        "CurrentReport = " & CurrentReport
End Sub

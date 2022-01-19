VERSION 5.00
Begin VB.Form frmWelcomeSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHello 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Timer tmrAutoClose 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   0
         Top             =   0
      End
      Begin VB.Label lblHello 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¡Hola, USERNAME!"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   405
         Left            =   0
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmWelcomeSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' Asignar valores
    CurrentTechnicianIndex = GetAssociatedTechnician(GetCurrentUserName)
    CurrentTechnician = ReadIni(Cfg.Data, "TECNICOS", CurrentTechnicianIndex, "<N/A>")
    SetFrmMainCaption
    
' Hola!
    picHello.Move 0, 0
    lblHello.Caption = "¡Hola, " & GetCurrentUserName(False) & "!"
    
    picHello.Visible = True
    
' Cerrar el form
    tmrAutoClose.Enabled = True
End Sub

Private Sub lblHello_Click()
    tmrAutoClose.Enabled = False
    Unload Me
End Sub

Private Sub picHello_Click()
    tmrAutoClose.Enabled = False
    Unload Me
End Sub

Private Sub tmrAutoClose_Timer()
    Unload Me
End Sub

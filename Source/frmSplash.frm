VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":65AA
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   1500
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblLoading 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblResolution 
      Alignment       =   2  'Center
      Caption         =   "480 x 272"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRenamon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "renamon-@live.com.ar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   5400
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      Top             =   3840
      Width           =   1755
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":C6A6
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1065
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   795
      Width           =   4815
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Pedidos Lite"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Form_Click()
    If fFromFrmMain Then
        Unload Me
        fFromFrmMain = False
        Exit Sub
    Else
        Call tmrSplash_Timer
    End If
End Sub

Private Sub Form_Initialize()
' Cargar estilo de controles de Windows mediante el uso del archivo manifest
    InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If fFromFrmMain Then
            Unload Me
            fFromFrmMain = False
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & " r" & App.Revision & " (Beta)"
    tmrSplash.Enabled = IIf(fFromFrmMain, False, True)
    lblLoading.Visible = IIf(fFromFrmMain, False, True)
End Sub

Private Sub Label_Click()
    Call Form_Click
End Sub

Private Sub lblMain_Click()
    Call Form_Click
End Sub

Private Sub lblRenamon_Click()
    NewMail Me, lblRenamon.Caption, " ", ""
End Sub

Private Sub lblVersion_Click()
    Call Form_Click
End Sub

Private Sub tmrSplash_Timer()
' Solo se ejecuta si el flag fFromFrmMain no esta activo

    If Not fFromFrmMain Then
    ' Cachea el último pedido e informe al inicio
        WriteIni Cfg.Link, "Last", cReport, LastReportFull
        WriteIni Cfg.Link, "Last", cTask, LastTaskFull
    End If

    Unload Me
End Sub

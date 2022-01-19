VERSION 5.00
Begin VB.Form frmPickTechnician 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar técnico"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPickTechnician.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraControls 
      Caption         =   "Seleccionar otro"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      Begin VB.ComboBox cmbTechnician 
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCurrentTechnician 
      Caption         =   "Técnico actual: XXX"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmPickTechnician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    With cmbTechnician
        If .ListIndex = -1 Then
            MsgBox "Debe seleccionar un técnico", vbExclamation, "Error"
            Exit Sub
        End If
        
        CurrentTechnician = .List(.ListIndex)
        CurrentTechnicianIndex = .ListIndex
    End With

    SetFrmMainCaption

    Unload Me
End Sub

Private Sub cmdCancel_Click()
' Si usa el selector de tecnico desde el menú de frmMain, no cerrar el programa
    If fFromFrmMain Then
       Unload Me
    Else
        Destroy True
    End If
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim i As Long

    lblCurrentTechnician.Caption = "Técnico actual: " & CurrentTechnician
    cmbTechnician.Clear
    
' Llenar la lista
    For i = 0 To TechnicianAmount
        cmbTechnician.AddItem ReadIni(Cfg.Data, "TECNICOS", i)
    Next
' Asignar el ulitmo tecnico logueado
    cmbTechnician.ListIndex = CurrentTechnicianIndex
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Destroy
        'End
    End If
End Sub

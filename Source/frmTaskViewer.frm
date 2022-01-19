VERSION 5.00
Begin VB.Form frmTaskViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmTaskViewer"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaskViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtMain 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mAddReport 
         Caption         =   "&Añadir informe"
         Shortcut        =   ^M
      End
      Begin VB.Menu mViewReports 
         Caption         =   "&Ver informes"
         Shortcut        =   ^I
      End
      Begin VB.Menu mEditTask 
         Caption         =   "&Editar pedido"
         Shortcut        =   ^E
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mCopyAsMailSubject 
         Caption         =   "Copiar como asunto de mail"
      End
      Begin VB.Menu mCopy 
         Caption         =   "Copiar pedido en el portapapeles"
      End
      Begin VB.Menu mSendByMail 
         Caption         =   "Enviar pedido por correo"
      End
      Begin VB.Menu mExport 
         Caption         =   "E&xportar pedido"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmTaskViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ApplyFont txtMain, fViewers

' Oh, no. Not again
    mViewReports.Visible = Not fFromFrmReportViewer
    mViewReports.Enabled = IIf(GetReportAmount(TaskHolder) > 0, True, False)
    
' Separamos los comandos para poder llamarlos desde fuera del form de ser necesario
    Update
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    fFromFrmTaskViewer = False
End Sub

Private Sub mAddReport_Click()
On Error Resume Next

    If GetStatus(TaskHolder) = strOK Then
        MsgBox "El pedido " & FormatTaskNumber(TaskHolder) & " está marcado como resuelto (Estado OK) y no se le pueden añadir más informes.", vbInformation, "Pedido cerrado"
        Exit Sub
    End If
    
    fFromFrmTaskViewer = True
    fNewReport = True
    frmNewReport.Show vbModal, Me
End Sub

Private Sub mCopy_Click()
    CopyToClipboard TaskHolder, True
End Sub

Private Sub mCopyAsMailSubject_Click()
    CopyAsSubject (TaskHolder)
End Sub

Private Sub mEditTask_Click()
    FromWhichForm (cFromTaskViewer)
    fNewTask = False
    
    frmNewTask.Show 1
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub mExport_Click()
    ExportTask (TaskHolder)
End Sub

Private Sub mSendByMail_Click()
    SendByMail TaskHolder, True, Me
End Sub

Private Sub mViewReports_Click()
    FromWhichForm (cFromTaskViewer)
    frmReportViewer.Show vbModal, Me
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub Update()
    txtMain.Text = DisplayTask(TaskHolder, True)
    Me.Caption = "Pedido " & FormatTaskNumber(TaskHolder)
End Sub

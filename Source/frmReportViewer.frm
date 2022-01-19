VERSION 5.00
Begin VB.Form frmReportViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmReportViewer"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewReport 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5370
      TabIndex        =   7
      Top             =   75
      Width           =   375
   End
   Begin VB.TextBox txtFocus 
      Height          =   360
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Renamon"
      Top             =   5640
      Width           =   315
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   75
      Width           =   375
   End
   Begin VB.CommandButton cmdRewind 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   75
      Width           =   375
   End
   Begin VB.Frame fraReport 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   5655
      Begin VB.TextBox txtReport 
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
         Height          =   5175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Label lblIndex 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3780
      TabIndex        =   6
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   3255
   End
   Begin VB.Menu mFile 
      Caption         =   "Archivo"
      Begin VB.Menu mAddReport 
         Caption         =   "&Añadir informe"
      End
      Begin VB.Menu mEditReport 
         Caption         =   "&Editar informe"
         Shortcut        =   ^E
      End
      Begin VB.Menu mViewAssociatedTask 
         Caption         =   "&Ver pedido asociado"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mCopy 
         Caption         =   "Copiar informe en el portapapeles"
      End
      Begin VB.Menu mSendByMail 
         Caption         =   "Enviar informe por correo"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aReports()  As Long     ' Contenedor de los informes de un pedido
Dim cReports    As Long     ' Cantidad de informes
Dim Index       As Long     ' Número de informe relativo al pedido (índice)
Dim RefTask     As Long     ' Almacena temporalmente el pedido de referencia

Private Sub cmdNewReport_Click()
    Call mAddReport_Click
End Sub

Private Sub cmdRewind_Click()
    If Index - 1 < 1 Then Exit Sub
    
    Index = Index - 1
    Display (Index)
End Sub

Private Sub cmdForward_Click()
    If Index + 1 > cReports Then Exit Sub
    
    Index = Index + 1
    Display (Index)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    fFromFrmReportViewer = False
End Sub

Private Sub mAddReport_Click()
    fFromFrmReportViewer = True
    fNewReport = True
    frmNewReport.Show vbModal, Me
End Sub

Private Sub mCopy_Click()
    CopyToClipboard CurrentReport, False
End Sub

Private Sub mEditReport_Click()
' El flag fNewReport permite abrir el formulario en modo editor. Al volver, mostrar el informe especifico
    FromWhichForm (cFromReportViewer)
    fNewReport = False
    fShowSpecificReport = True
     
    frmNewReport.Show 1
End Sub

Private Sub Form_Load()
    ApplyFont txtReport, fViewers
    Call Update
    mViewAssociatedTask.Visible = Not fFromFrmTaskViewer
    
    If GetReportAmount(TaskHolder) > 0 Then
        mCopy.Enabled = True
        mSendByMail.Enabled = True
    Else
        mCopy.Enabled = False
        mSendByMail.Enabled = False
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
            Exit Sub
              
        Case vbKeyLeft
              Call cmdRewind_Click
              txtFocus.SetFocus
              Exit Sub
              
        Case vbKeyRight
            Call cmdForward_Click
            txtFocus.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub mSendByMail_Click()
    SendByMail CurrentReport, False, Me
End Sub

Private Sub mViewAssociatedTask_Click()
'   fSearchByTask = True
    fFromFrmReportViewer = True
    frmTaskViewer.Show vbModal, Me
End Sub

Sub Update()
' Traer los informes (si los hay)
    If fFromFrmSearch Then
        If Not fFromFrmTaskViewer Then
            TaskHolder = GetReferenceTask(CurrentReport)
        End If
    End If
    
    If fDebug Then
        Log "Executing frmReportViewer.Update, fShowSpecificReport flag is " & fShowSpecificReport & vbNewLine & _
            "Current Report is " & CurrentReport & vbNewLine & _
            "Current Task is " & TaskHolder
    End If
    
    aReports = GetTaskReports(TaskHolder)
    cReports = UBound(aReports)
    Index = cReports
    Me.Caption = "Informes del pedido " & FormatTaskNumber(TaskHolder)
    
' Elegir el modo de mostrar el informe
    If fShowSpecificReport Then
        SetReportQueue mSpecific, CurrentReport
    Else
        SetReportQueue mAll
    End If
    
    If UCase(GetStatus(TaskHolder)) = strOK Then
        cmdNewReport.Enabled = False
        mAddReport.Enabled = False
    Else
        cmdNewReport.Enabled = True
        mAddReport.Enabled = True
    End If
End Sub

Sub SetReportQueue(ByVal Mode As Byte, Optional ByVal SpecificReport As Long)
Dim i As Long

    Select Case Mode
        Case mAll
            Display (Index)
          
        Case mSpecific
            For i = 1 To UBound(aReports)
                If aReports(i) = SpecificReport Then
                    Index = i
                    Exit For
                End If
            Next i
          
            Display (Index)
    End Select
End Sub

Sub Display(Index As Long)
Dim tDate As String

' Habilitar / Deshabilitar botones
    If cReports < 2 Then
        cmdRewind.Enabled = False
        cmdForward.Enabled = False
        
        If cReports = 0 Then
            lblIndex.Caption = ""
            lblMain.Caption = ""
            mEditReport.Enabled = False
            txtReport.Text = "<El pedido " & FormatTaskNumber(TaskHolder) & " no tiene informes.>"
            Exit Sub
        End If
    Else
        If Index + 1 > cReports Then
            cmdRewind.Enabled = True
            cmdForward.Enabled = False
        ElseIf Index - 1 < 1 Then
            cmdRewind.Enabled = False
            cmdForward.Enabled = True
        Else
            cmdRewind.Enabled = True
            cmdForward.Enabled = True
        End If
    End If

' Mostrar el informe Index
    CurrentReport = aReports(Index)
    TaskHolder = GetReferenceTask(CurrentReport)
    tDate = GetDate(CurrentReport, False)
    
' Procesar textos
    lblMain.Caption = "Informe " & FormatReportNumber(aReports(Index)) & "  |  " & Mid$(tDate, 1, 2) & "/" & Mid$(tDate, 3, 2) & "/" & Mid$(tDate, 5, 2)
    txtReport.Text = DisplayReport(CurrentReport)
    lblIndex.Caption = Index & " / " & cReports
    Me.Caption = "Informe " & FormatReportNumber(aReports(Index)) & " del pedido " & FormatTaskNumber(TaskHolder)
End Sub

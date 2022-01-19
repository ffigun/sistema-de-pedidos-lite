VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personalización"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optMailAndStuff 
      Caption         =   "Mail y Otros"
      Height          =   285
      Left            =   4920
      TabIndex        =   40
      Top             =   120
      Width           =   2040
   End
   Begin VB.OptionButton optMain 
      Caption         =   "Opciones"
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton optFont 
      Caption         =   "Fuentes"
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   120
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraMailAndStuff 
      Caption         =   "Mail y otros"
      Height          =   3495
      Left            =   120
      TabIndex        =   31
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtClipboardLen 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   3
         TabIndex        =   39
         Text            =   "80"
         ToolTipText     =   $"frmOptions.frx":0000
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkMsgBox 
         Caption         =   "Notificar"
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
         Left            =   2280
         TabIndex        =   34
         ToolTipText     =   "Mostrar un mensaje en pantalla cada vez que se copia satisfactoriamente el contenido al portapapeles."
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtWikiPath 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Indica la ruta absoluta del ejecutable de la Wiki. No hay comodines disponibles."
         Top             =   3000
         Width           =   6615
      End
      Begin VB.TextBox txtSubjectFormat 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   32
         Text            =   "%N - %D"
         ToolTipText     =   "Puede utilizar estos comodines: %K = Tipo (Informe / Pedido), %N = Número, %D = Detalle, %C = Contacto, %T = Técnico"
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label7 
         Caption         =   "caracteres"
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
         Left            =   720
         TabIndex        =   38
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Longitud máxima del asunto:"
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
         TabIndex        =   37
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Ruta de la Wiki:"
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
         TabIndex        =   36
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Formato del asunto:"
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
         TabIndex        =   35
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Opciones"
      Height          =   3495
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtSepAmount 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "3"
         ToolTipText     =   $"frmOptions.frx":008A
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "Mostrar todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "PEND"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkHightlightWithReports 
         Caption         =   "Destacar con (*) los pedidos con informes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Muestra un * al lado de los pedidos que tienen informes."
         Top             =   3120
         Width           =   6615
      End
      Begin VB.CheckBox chkShowPriority 
         Caption         =   "Mostrar prioridad en la lista principal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Muestra el valor numérico que representa la prioridad al lado de los informes."
         Top             =   2760
         Width           =   6615
      End
      Begin VB.Label Label5 
         Caption         =   "espacios"
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
         Left            =   720
         TabIndex        =   30
         Top             =   2190
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Separación de campos de la lista principal:"
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
         TabIndex        =   28
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Filtrar pedidos por estado:"
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
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fraFont 
      Caption         =   "Fuentes"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CheckBox chkUseGlobalFont 
         Caption         =   "Utilizar una única fuente:"
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
         TabIndex        =   17
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtFont 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   3240
         TabIndex        =   16
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtSize 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtFont 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   3240
         TabIndex        =   8
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtFont 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtFont 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtFont 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblWhy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Por qué no puedo cambiar el tamaño de algunas fuentes?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   105
         MousePointer    =   10  'Up Arrow
         TabIndex        =   18
         Top             =   3120
         Width           =   4845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   6600
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label lblCaption 
         Caption         =   "Fuente del Panel de búsqueda:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lblCaption 
         Caption         =   "Fuente de Visores:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblCaption 
         Caption         =   "Fuente del Panel lateral:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblCaption 
         Caption         =   "Fuente de Lista principal:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseGlobalFont_Click()
    If chkUseGlobalFont.Value = vbChecked Then
        EnableButtons False
        EnableButtons True, True
    Else
        EnableButtons True
        EnableButtons False, True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Integer

' Validar
    For i = 0 To 4
        If Not IsNumeric(txtSize(i).Text) Then
            fNotNumeric = True
            Exit For
        End If
    Next i

    If fNotNumeric Then
        MsgBox "Introduzca sólo valores numéricos en los campos de tamaño de fuente.", vbExclamation
        fNotNumeric = False
        Exit Sub
    End If
    
    If Not IsNumeric(txtClipboardLen.Text) Then
        MsgBox "Introduzca sólo valores numéricos en el campo de caracteres a copiar en el portapapeles.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtSepAmount) Then
        MsgBox "Introduzca sólo valores numéricos en el campo de separación de campos.", vbExclamation
        Exit Sub
    End If
    
    For i = 0 To 4
        sFontName(i) = txtFont(i).Text
        sFontSize(i) = Val(txtSize(i).Text)
    Next i
    
    fGlobalFont = IIf(chkUseGlobalFont.Value = vbChecked, True, False)
    
' Guardar y aplicar a frmMain, los demás se actualizan con sus _Load correspondientes
    For i = 0 To 4
        WriteIni Cfg.Data, "FONT", "Fuente" & i, sFontName(i)
        WriteIni Cfg.Data, "FONT", "Grado" & i, sFontSize(i)
    Next i

    WriteIni Cfg.Data, "FONT", "FuenteUnica", IIf(fGlobalFont, "1", "0")
    
    ApplyFont frmMain.lstData, fMain
    ApplyFont frmMain.txtTask, fSidePanel
    
' Opciones
    SubjectLen = txtClipboardLen.Text
    WriteIni Cfg.Data, "EXTRA", "SubjectLen", txtClipboardLen.Text
     
    fClipboardCopyNotify = chkMsgBox.Value
    WriteIni Cfg.Data, "EXTRA", "ClpbrdNotify", IIf(chkMsgBox.Value, 1, 0)
     
    SubjectFormat = txtSubjectFormat.Text
    WriteIni Cfg.Data, "EXTRA", "SubjectFormat", txtSubjectFormat.Text
     
    gSepChar = txtSepAmount.Text
    WriteIni Cfg.Data, "FONT", "gSepChar", txtSepAmount.Text
     
    fShowPriority = chkShowPriority.Value
    WriteIni Cfg.Data, "FLAGS", "MostrarPrioridad", IIf(chkShowPriority.Value, 1, 0)
     
    fHlTasksWithReports = chkHightlightWithReports.Value
    WriteIni Cfg.Data, "FLAGS", "DestacarPedidosConInformes", IIf(chkHightlightWithReports.Value, 1, 0)
     
    If optFilter(cPEND).Value Then MainFilter = cPEND
    If optFilter(cOK).Value Then MainFilter = cOK
    If optFilter(cTODOS).Value Then MainFilter = cTODOS
     
    WriteIni Cfg.Data, "FLAGS", "Mostrar", MainFilter
     
    WikiPath = txtWikiPath.Text
    WriteIni Cfg.Data, "FGWIKI", "RutaExe", txtWikiPath.Text
     
    frmMain.AddTask (MainFilter)
     
    If fDebug Then
    Log "It was raised the event _Click of cmdOk in frmOptions." & vbNewLine & _
        "Options were set."
    End If
     
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Si aprieta Ctrl + Enter
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        Call cmdOk_Click
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    
    Me.Icon = frmMain.Icon
    
    For i = 0 To 4
        txtFont(i).Text = sFontName(i)
        txtSize(i).Text = Trim(Str(sFontSize(i)))
    Next i
    
' Actualizar y validar
    chkUseGlobalFont.Value = ReadIni(Cfg.Data, "FONT", "FuenteUnica", "0")
    Call chkUseGlobalFont_Click
    
    optFilter(MainFilter).Value = True
    chkShowPriority.Value = IIf(fShowPriority, vbChecked, vbUnchecked)
    chkHightlightWithReports.Value = IIf(fHlTasksWithReports, vbChecked, vbUnchecked)
    txtWikiPath.Text = WikiPath
    txtSepAmount.Text = gSepChar
    txtClipboardLen.Text = SubjectLen
    chkMsgBox.Value = IIf(fClipboardCopyNotify, 1, 0)
    txtSubjectFormat.Text = SubjectFormat

    Call optMain_Click
End Sub

Private Sub lblWhy_Click()
    MsgBox "Este software no permite modificar el tamaño de las fuentes del panel lateral y de los visores de pedidos e informes." & vbNewLine & vbNewLine & _
           "Esto se debe a la forma en que el software procesa los datos que se muestran en pantalla. Para mejores resultados de legibilidad, se recomienda utilizar fuentes de ancho fijo (Monospaced fonts) tales como Courier, Courier New, Consolas, o Lucida Sans Typewriter.", vbInformation, "Ayuda"
End Sub

Private Sub optMailAndStuff_Click()
    fraFont.Enabled = False
    fraFont.Visible = False
     
    fraOptions.Enabled = False
    fraOptions.Visible = False
     
    fraMailAndStuff.Enabled = True
    fraMailAndStuff.Visible = True
End Sub

Private Sub optMain_Click()
    fraFont.Enabled = False
    fraFont.Visible = False
     
    fraOptions.Enabled = True
    fraOptions.Visible = True
     
    fraMailAndStuff.Enabled = False
    fraMailAndStuff.Visible = False
End Sub

Private Sub optFont_Click()
    fraOptions.Enabled = False
    fraOptions.Visible = False
     
    fraFont.Enabled = True
    fraFont.Visible = True
     
    fraMailAndStuff.Enabled = False
    fraMailAndStuff.Visible = False
End Sub

Sub EnableButtons(ByVal EnableObject As Boolean, Optional ByVal EnableGlobalFont As Boolean = False)
Dim i As Integer

If Not EnableGlobalFont Then
    For i = 0 To 3
        txtFont(i).Enabled = EnableObject
        txtSize(i).Enabled = EnableObject
         
        txtFont(i).BackColor = IIf(EnableObject, vbWindowBackground, vbButtonFace)
        txtSize(i).BackColor = IIf(EnableObject, vbWindowBackground, vbButtonFace)
         
        txtFont(i).ForeColor = IIf(EnableObject, vbButtonText, vbGrayText)
        txtSize(i).ForeColor = IIf(EnableObject, vbButtonText, vbGrayText)
         
        lblCaption(i).ForeColor = IIf(EnableObject, vbButtonText, vbGrayText)
    Next i
    
Else
    txtFont(fGlobal).Enabled = EnableObject
    txtSize(fGlobal).Enabled = EnableObject
     
    txtFont(fGlobal).BackColor = IIf(EnableObject, vbWindowBackground, vbButtonFace)
    txtSize(fGlobal).BackColor = IIf(EnableObject, vbWindowBackground, vbButtonFace)
End If
    
    ' --- Configuración forzada ---
    ' Para mantener la estructura de la interfaz
        txtSize(fViewers).Enabled = False
        txtSize(fViewers).ForeColor = vbGrayText
        txtSize(fViewers).BackColor = vbButtonFace
         
        txtSize(fSidePanel).Enabled = False
        txtSize(fSidePanel).ForeColor = vbGrayText
        txtSize(fSidePanel).BackColor = vbButtonFace
    ' --- Fin de configuración forzada ---
End Sub

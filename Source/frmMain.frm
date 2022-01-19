VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   12990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4980
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":038A
      Left            =   120
      List            =   "frmMain.frx":038C
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Frame fraTask 
      Height          =   5175
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.TextBox txtTask 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmMain.frx":038E
         Top             =   240
         Width           =   4515
      End
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Left            =   9960
      Picture         =   "frmMain.frx":03B9
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgError 
      Height          =   240
      Left            =   9960
      Picture         =   "frmMain.frx":0743
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mNew 
         Caption         =   "&Nuevo"
         Begin VB.Menu mNewTask 
            Caption         =   "&Pedido"
            Shortcut        =   ^N
         End
         Begin VB.Menu mNewReport 
            Caption         =   "&Informe"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu mExport 
         Caption         =   "E&xportar"
         Begin VB.Menu mExportSelectedTask 
            Caption         =   "&Pedido Seleccionado"
            Shortcut        =   ^A
         End
         Begin VB.Menu mSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mExportTasks 
            Caption         =   "&Todos los Pedidos"
         End
         Begin VB.Menu mExportReports 
            Caption         =   "Todos los &Informes"
         End
      End
      Begin VB.Menu mEditTask 
         Caption         =   "&Editar pedido"
         Shortcut        =   ^E
      End
      Begin VB.Menu mSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mCopyToClipboard 
         Caption         =   "&Copiar pedido en el portapapeles"
      End
      Begin VB.Menu mCopyAsSubject 
         Caption         =   "C&opiar como asunto de mail"
      End
      Begin VB.Menu mSendByMail 
         Caption         =   "E&nviar pedido por correo"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&Ver"
      Begin VB.Menu mViewTask 
         Caption         =   "&Pedido"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mViewReports 
         Caption         =   "&Informes"
         Shortcut        =   ^I
      End
      Begin VB.Menu mViewOkReport 
         Caption         =   "Informe de &cierre"
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Opciones"
      Begin VB.Menu mSearch 
         Caption         =   "&Buscar..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mWiki 
         Caption         =   "&Abrir Wiki"
      End
      Begin VB.Menu mCustomize 
         Caption         =   "Personalizar"
      End
      Begin VB.Menu mSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mTechnicianSelect 
         Caption         =   "&Seleccionar técnico (Actual: XXX)"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Begin VB.Menu popNewReport 
         Caption         =   "Nuevo informe"
      End
      Begin VB.Menu popNewTask 
         Caption         =   "Nuevo pedido"
      End
      Begin VB.Menu popSep1 
         Caption         =   "-"
      End
      Begin VB.Menu popViewTask 
         Caption         =   "Ver Pedido"
      End
      Begin VB.Menu popEditTask 
         Caption         =   "Editar pedido"
      End
      Begin VB.Menu popViewReports 
         Caption         =   "Ver informes"
      End
      Begin VB.Menu popViewOkReport 
         Caption         =   "Ver informe de cierre"
      End
      Begin VB.Menu popSep2 
         Caption         =   "-"
      End
      Begin VB.Menu popCopySubject 
         Caption         =   "Copiar como asunto de mail"
      End
      Begin VB.Menu popCopyToClipboard 
         Caption         =   "Copiar pedido en el portapapeles"
      End
      Begin VB.Menu popSendByMail 
         Caption         =   "Enviar pedido por correo"
      End
      Begin VB.Menu popExportTask 
         Caption         =   "Exportar pedido"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&?"
      Begin VB.Menu mAbout 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu mSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mDebug 
         Caption         =   "&Funciones experimentales"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurListIndex As Long       ' Indice actual de la lista, para evitar un flood de eventos

Private Sub Form_Load()
Dim iFont As Integer

    ' En frmSplash se cachean los valores de LastTask y LastReport!
    frmSplash.Show vbModal, Me
    
    CheckThreshold

    Me.Width = frmMainW
    Me.Height = frmMainH

    mblnRightClick = False
    mnuPopUp.Visible = False
    
    CurListIndex = -1
    
    fHlTasksWithReports = IIf(Val(ReadIni(Cfg.Data, "FLAGS", "DestacarPedidosConInformes", 0)) = 0, False, True)
    fShowPriority = IIf(Val(ReadIni(Cfg.Data, "FLAGS", "MostrarPrioridad", 0)) = 0, False, True)
    fShowWiki = IIf(Val(ReadIni(Cfg.Data, "FGWIKI", "Mostrar", 0)) = 0, False, True)
    
    WikiPath = RTrim(LTrim(ReadIni(Cfg.Data, "FGWIKI", "RutaExe", "")))
    MainFilter = CByte(ReadIni(Cfg.Data, "FLAGS", "Mostrar", 2))
    CurrentTechnicianIndex = CLng(ReadIni(Cfg.Data, "TECNICOS", "UltimoLogueado", "0"))
    CurrentTechnician = ReadIni(Cfg.Data, "TECNICOS", CurrentTechnicianIndex, "<N/A>")
    
    If fDebug Then
        Log "At _Load event of frmMain. Welcome."
    End If
    
' Si no hay ruta de Wiki, no usar aunque lo especifiquen las configuraciones
    If Trim(WikiPath) = "" Then
        fShowWiki = False
    End If
    
    Me.Caption = "Sistema de pedidos lite (Beta) | " & App.Major & "." & App.Minor & " r" & App.Revision & " | " & "Técnico actual: " & CurrentTechnician
    
    lstData.Clear
    
    txtTask.Text = EmptyDetails
    mTechnicianSelect.Caption = "Seleccionar técnico (Actual: " & CurrentTechnician & ")"

    AddTask (MainFilter)
    mWiki.Visible = fShowWiki

    ApplyFont lstData, fMain
    ApplyFont txtTask, fSidePanel

' Dar a elegir un tecnico o mostrar la pantalla de bienvenida si reconoce el nombre de usuario de Windows
    If GetAssociatedTechnician(GetCurrentUserName) <> -1 Then
        frmWelcomeSplash.Show vbModal, Me
    Else
        frmPickTechnician.Show vbModal, Me
    End If
   
   Me.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Destroy
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fDebug Then
    Log "QueryOnload event raised on frmMain." & vbNewLine & _
        "Cancel is " & Cancel & vbNewLine & _
        "UnloadMode is " & UnloadMode
    End If
    
    If UnloadMode = vbFormCode Then Exit Sub
     
    If Not Destroy Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next

' Esquivando errores con mi bici
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If

' Establecer W y H en variables, cuando cierren el programa se guardan estos datos
    CheckThreshold
    
' Constructor de interfaz
    With fraTask
        .Left = Me.ScaleWidth - .Width - wOffset * 0.4
        .Height = Me.ScaleHeight - 100
        
        lstData.Width = .Left - wOffset
        lstData.Height = Me.ScaleHeight - hOffset
        txtTask.Height = .Height - (hOffset * 1.5)
        txtTask.Width = 4515
    End With
End Sub

Private Sub lstData_Click()
' Este codigo NO SE PUEDE PROBAR EN EL IDE DE VB. SOLO SE VEN LOS RESULTADOS UNA VEZ COMPILADO.
' Debe tener que ver con cómo se hacen los hooks de las API de mouse dentro del IDE de VB.

On Error Resume Next
' Si el mouse no está en una parte en blanco
    If Not ItemHover = -1 Then
        TogglePopMenuTaskOptions True
    Else
        TogglePopMenuTaskOptions False
    End If
 
' Cargar sólo si hay algo seleccionado
    If lstData.ListIndex <> -1 Then
        If lstData.ListIndex <> CurListIndex Then
           SetCurrentTask
           txtTask.Text = DisplayTask(CurrentTask) ' IIf(fDisplayTaskAndReport, DisplayTask2(CurrentTask), DisplayTask(CurrentTask))
        End If
    End If

' Si apretó el clic derecho
    If mblnRightClick Then
        If Not ItemHover = -1 Then
            If GetStatus(CurrentTask) = strOK Then
                popNewReport.Visible = False
                popViewOkReport.Visible = True
            Else
                popNewReport.Visible = True
                popViewOkReport.Visible = False
            End If
        End If
        
            Me.PopupMenu mnuPopUp
    
        ' "Soltar" click derecho virtual
        mblnRightClick = False
    End If
  
' Recuerda el ultimo presionado
    CurListIndex = lstData.ListIndex
End Sub

Sub TogglePopMenuTaskOptions(ByVal bValue As Boolean)
    popNewReport.Visible = bValue
    popViewOkReport.Visible = bValue
    popEditTask.Visible = bValue
    popViewReports.Visible = bValue
    popViewTask.Visible = bValue
    popSep1.Visible = bValue
    popCopySubject.Visible = bValue
    popSendByMail.Visible = bValue
    popCopyToClipboard.Visible = bValue
    popExportTask.Visible = bValue
    popSep2.Visible = bValue
End Sub

Private Sub lstData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift <> vbCtrlMask Then
        If lstData.ListIndex = -1 Then Exit Sub
        If lstData.ListCount = 0 Then Exit Sub
        
        Call mViewReports_Click
    End If

    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        If lstData.ListIndex = -1 Then Exit Sub
        If lstData.ListCount = 0 Then Exit Sub

        If UCase(GetStatus(CurrentTask)) = strOK Then
            Call mViewReports_Click
        Else
            Call mNewReport_Click
        End If
    End If
End Sub

Private Sub lstData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Seleccionar el item al hacer clic derecho simulando haber presionado clic izquierdo
Dim lItem As Long
    If Button = vbRightButton Then
        mouse_event MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, X, Y, 0, 0
    
        mblnRightClick = True
    End If
End Sub

Private Sub lstData_DblClick()
' Esto no se puede probar desde el IDE, solo se ve desde la app compilada
    If ItemHover = -1 Then
        Call mNewTask_Click
        Exit Sub
    End If

    If lstData.ListIndex = -1 Then Exit Sub
    If lstData.ListCount = 0 Then Exit Sub
      
    Call mViewReports_Click
End Sub

Private Sub lstData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Tener siempre conocimiento del item que esta debajo del mouse
    ItemHover = ItemUnderMouse(lstData.hwnd, X, Y)
End Sub

Private Sub mCopy_Click()
    Clipboard.Clear
    Clipboard.SetText DisplayReport(CurrentReport, True)
End Sub

Private Sub mCopyAsSubject_Click()
    CopyAsSubject (CurrentTask)
End Sub

Private Sub mDebug_Click()
    frmBeta.Show vbModeless, Me
End Sub

Private Sub mFile_Click()
    If lstData.ListIndex = -1 Or lstData.ListCount = 0 Then
        mEditTask.Enabled = False
        mCopyToClipboard.Enabled = False
        mCopyAsSubject.Enabled = False
        mSendByMail.Enabled = False
        mExportSelectedTask.Enabled = False
        mNewReport.Enabled = False
   Else
        mEditTask.Enabled = True
        mCopyToClipboard.Enabled = True
        mCopyAsSubject.Enabled = True
        mSendByMail.Enabled = True
        mExportSelectedTask.Enabled = True
        mNewReport.Enabled = IIf(GetStatus(CurrentTask) = strOK, False, True)
    End If
End Sub

Private Sub mAbout_Click()
' Llamar al SplashScreen sin el timer que cierra sóla la ventana
    fFromFrmMain = True
    frmSplash.Show vbModal, Me
End Sub

Private Sub mEditTask_Click()
    SetTaskHolder
    fNewTask = False
    frmNewTask.Show vbModal, Me
End Sub

Private Sub mExportReports_Click()
    Export (cReports)
End Sub

Private Sub mExportTasks_Click()
    Export (cTasks)
End Sub

Private Sub mExportSelectedTask_Click()
    ExportTask (CurrentTask)
End Sub

Private Sub mCustomize_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mSearch_Click()
' Guardar el valor actual de CurrentTask para traerlo luego
    SetTaskHolder
    frmSearch.Show vbModal, Me
End Sub

Private Sub mNewReport_Click()
On Error Resume Next

    SetTaskHolder
    fNewReport = True
    
    If GetStatus(CurrentTask) = strOK Then
        MsgBox "El pedido " & FormatTaskNumber(CurrentTask) & " está marcado como resuelto (Estado OK) y no se le pueden añadir más informes.", vbInformation, "Pedido cerrado"
        Exit Sub
    End If
    
    frmNewReport.Show vbModal, Me
End Sub

Private Sub mNewTask_Click()
    fNewTask = True
    frmNewTask.Show vbModal, Me
End Sub

Private Sub mExit_Click()
    Destroy
End Sub

Private Sub mTechnicianSelect_Click()
    fFromFrmMain = True
    frmPickTechnician.Show vbModal, Me
End Sub

Private Sub mSendByMail_Click()
    SendByMail CurrentTask, True, Me
End Sub

Private Sub mView_Click()
    If lstData.ListIndex = -1 Then
        mViewReports.Enabled = False
        mViewOkReport.Enabled = False
        mViewTask.Enabled = False
    Else
        mViewReports.Enabled = True
        mViewOkReport.Enabled = True
        mViewTask.Enabled = True
        
        If GetStatus(CurrentTask) = strPEND Then
            mViewOkReport.Enabled = False
        End If
    End If
End Sub

Private Sub mViewOkReport_Click()
Dim iTemp As Long
    SetTaskHolder

    iTemp = CLng(GetOkReport((CurrentTask)))

    If iTemp = -1 Then Exit Sub

' Mostrar el informe especifico (aunque técnicamente es el último cargado)
    fShowSpecificReport = True
    CurrentReport = iTemp
    frmReportViewer.Show vbModal, Me
End Sub

Private Sub mViewReports_Click()
    SetTaskHolder
    frmReportViewer.Show vbModal, Me
End Sub

Private Sub mViewTask_Click()
    Call popViewTask_Click
End Sub

Private Sub mWiki_Click()
On Error GoTo OhNo
    Shell WikiPath, vbNormalFocus
    Exit Sub

OhNo:
    MsgBox "No se pudo iniciar el programa. Compruebe que la ruta absoluta especificada en el archivo de configuración sea correcta, luego compruebe que el programa exista y que tenga acceso a él.", vbExclamation, "Error al abrir la Wiki"
End Sub

Private Sub popCopySubject_Click()
    CopyAsSubject (CurrentTask)
End Sub

Private Sub popSendByMail_Click()
    SendByMail CurrentTask, True, Me
End Sub

Private Sub popCopyToClipboard_Click()
    CopyToClipboard CurrentTask, True
End Sub

Private Sub popExportTask_Click()
    Call mExportSelectedTask_Click
End Sub

Private Sub popViewOkReport_Click()
    Call mViewOkReport_Click
End Sub

Private Sub popNewReport_Click()
    Call mNewReport_Click
End Sub

Private Sub popEditTask_Click()
    Call mEditTask_Click
End Sub

Private Sub popNewTask_Click()
    Call mNewTask_Click
End Sub

Private Sub popViewReports_Click()
    Call mViewReports_Click
End Sub

Private Sub popViewTask_Click()
    SetTaskHolder
    frmTaskViewer.Show vbModal, Me
End Sub

Private Sub txtTask_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Destroy
End Sub

Sub CheckThreshold()
' WHThreshold es el porcentaje maximo de cHeight y cWidth.
' Si se superan estos valores, no se guardarán en el INI de Configuración y se volverá al tamaño default
    If frmMainH > Screen.Height * WHThreshold Or frmMainH < 7000 Then
        frmMainH = 7200
    End If
    
    If frmMainW > Screen.Width * WHThreshold Or frmMainW < 13000 Then
        frmMainW = 13500
    End If
End Sub

Function NormalizeWidth() As Long
    If Me.Width > Screen.Width * WHThreshold Or Me.Width < 13000 Then
        NormalizeWidth = 13500
    Else
        NormalizeWidth = Me.Width
    End If
End Function

Function NormalizeHeight() As Long
    If Me.Height > Screen.Height * WHThreshold Or Me.Height < 7000 Then
        NormalizeHeight = 7200
    Else
        NormalizeHeight = Me.Height
    End If
End Function

Function EmptyDetails() As String
Dim Buffer  As String
Dim pd      As Long:     pd = 30
Dim nl      As String:   nl = vbNewLine
Dim tl      As String:   tl = nl + nl + nl + nl

    Buffer = "Pedido: " & nl & _
             "Informes: " & nl & _
                            nl & _
             "Contacto: " & nl & _
             "Sucursal: " & nl & _
                            nl & _
             "Fecha: " & nl & _
             "Hora: " & nl & _
             "Fecha Ok: " & nl & _
                            nl & _
             "Estado: " & nl & _
             "Prioridad: " & nl & _
             "Técnico: " & nl & _
                           nl & _
      PadStr("--- Observaciones ", MaxChrW, "-") & tl & _
      PadStr("--- Detalle ", MaxChrW, "-")

    EmptyDetails = Buffer
End Function

Sub AddTask(ByVal bFilter As Byte)

'If LastTask = -1 Then
'    Exit Sub
'End If

Dim strA As String  ' Pedido formateado
Dim strB As String  ' Estado del pedido
Dim strC As String  ' Tiene informes
Dim strD As String  ' Prioridad
Dim strE As String  ' Detalle
Dim sTmp As String  ' Estado temporal
Dim Buff As String  ' Buffer
Dim lCnt As Long    ' Contador
Dim GoOn As Boolean ' Agregar pedido

' Inicializar
GoOn = True
lstData.Clear
txtTask.Text = EmptyDetails
Screen.MousePointer = vbHourglass

' Leer todos los pedidos y agregar uno a uno los items al buffer
For lCnt = 0 To LastTask

    sTmp = GetStatus(lCnt)

    If bFilter = cOK And sTmp <> strOK Then GoOn = False
    If bFilter = cPEND And sTmp <> strPEND Then GoOn = False
    If bFilter = cTODOS Then GoOn = True
    
    If GoOn Then
    ' Pedido
        strA = FormatTaskNumber(lCnt) & Space(gSepChar)
        
    ' Estado
        strB = PadStr(sTmp, 4) & Space(gSepChar)
        
    ' Tiene informes
        If fHlTasksWithReports Then
            If HasReports(lCnt) Then
                strC = "*" & Space(gSepChar)
            Else
                strC = Space(1) & Space(gSepChar)
            End If
        End If
        
    ' Prioridad
        If fShowPriority Then
            strD = GetPriority(lCnt) & Space(gSepChar)
        End If
        
    ' Detalle
        strE = Replace(GetDetail(lCnt, True), cNewLine, Space(1))
    
    ' Compilar datos y agregar
        lstData.AddItem strA & strB & strC & strD & strE
    End If
    
    GoOn = True
Next

    Screen.MousePointer = vbDefault
    
    If Not fFromFrmSearch Then SelectTaskByNumber (CurrentTask)
End Sub

Sub Update()
' No comprobar indices de list, ya que SetCurrentTask y SelectTaskByNumber pondran -1 a CurrentTask si algo falla
    SetCurrentTask
    SelectTaskByNumber (CurrentTask)
       
        If CurrentTask = -1 Then Exit Sub
       
    txtTask.Text = IIf(TaskExists(CurrentTask), DisplayTask(CurrentTask), EmptyDetails)
End Sub

Sub SetCurrentTask()
' Recuperar el valor anterior del pedido actual seleccionado en la lista principal de frmMain

' Si edito algo
    If fFromFrmTaskViewer Or fFromFrmReportViewer Then
        CurrentTask = TaskHolder
            If fDebug Then
                Log "Se asigno " & TaskHolder & " a CurrentTask porque fFromFrmTaskViewer o fFromFrmReportViewer estaban en 1"
            End If
        Exit Sub
    End If

' Bug-free
    If lstData.ListCount = 0 Or lstData.ListIndex = -1 Then
        CurrentTask = -1
            If fDebug Then
                Log "Se asigno -1 a CurrentTask porque .ListCount era 0 o .ListIndex era -1"
            End If
        Exit Sub
    End If

' Si cargo un nuevo pedido
    If fFromFrmNewTask Then
        CurrentTask = LastTask
        fFromFrmNewTask = False
            If fDebug Then
                Log "Se asigno " & LastTask & " a CurrentTask porque el usuario cargo un nuevo pedido"
            End If
        Exit Sub
    End If

' Sino, lo que este seleccionado
    CurrentTask = Mid$(lstData.List(lstData.ListIndex), 1, 5)
        If fDebug Then
            Log "Se asigno " & CurrentTask & " a CurrentTask porque no se cumplieron el resto de las condiciones"
        End If
End Sub

Sub SetTaskHolder()
' Permite que CurrentTask sea solo de frmMain y este indice no se modifique cuando la lista de pedidos de frmMain se actualiza
    TaskHolder = CurrentTask
End Sub

Sub SelectTaskByNumber(ByVal Number As Long)
Dim sT  As String
Dim i   As Long
Dim bOk As Boolean

sT = FormatTaskNumber(Number)

For i = 0 To lstData.ListCount - 1
    If Left$(lstData.List(i), 5) = sT Then
        lstData.ListIndex = i
        bOk = True
        Exit For
    End If
Next i

' Si sale del For y no se encontro el pedido...
    If Not bOk Then
        CurrentTask = -1
        txtTask.Text = EmptyDetails
            
        If fDebug Then
            Log "Se asigno -1 a CurrentTask porque luego de recorrer la lista no se encontro el pedido " & Number
        End If
    End If
End Sub

Attribute VB_Name = "modFunctions"
Option Explicit

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Declare Function ClientToScreen Lib "USER32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function LBItemFromPt Lib "COMCTL32.DLL" (ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, ByVal bAutoScroll As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const SW_SHOW = 5
Public Const MOUSEEVENTF_LEFTDOWN = &H2     ' Clic izquierdo virtual
Public Const MOUSEEVENTF_LEFTUP = &H4       ' Soltar clic izquierdo virtual

Public mblnRightClick As Boolean

Sub Main()
On Error GoTo SomethingWentWrong

' Acá comienza el programa
    Dim bTasksReportsOrLinkFailed As Boolean
    Dim sWhichFailed As String
    Cfg.Paths = App.Path & "\Rutas.ini"
    bTasksReportsOrLinkFailed = False

' Si un archivo no existe, notificar y abortar
    If Not PathExists(Cfg.Paths) Then
        MsgBox "El archivo de rutas no existe o no se puede tener acceso a él." & vbNewLine & vbNewLine & _
               "Se creará uno nuevo en la ruta del programa, modifique las rutas de los archivos en éste antes de volver a ejecutar la aplicación." _
               , vbExclamation, "No se encuentra el archivo"
            Call CreatePathFile
            End
    End If
    
    ' Si no existen archivos, crearlos o ignorarlos
    Cfg.Data = Replace(ReadIni(Cfg.Paths, "MAIN", "Datos", cAppPath & "\Datos.ini"), cAppPath, App.Path)
    Cfg.Link = Replace(ReadIni(Cfg.Paths, "MAIN", "Enlace", cAppPath & "\Enlace.ini"), cAppPath, App.Path)
    Cfg.Tasks = Replace(ReadIni(Cfg.Paths, "MAIN", "Pedidos", cAppPath & "\Pedidos.ini"), cAppPath, App.Path)
    Cfg.Reports = Replace(ReadIni(Cfg.Paths, "MAIN", "Informes", cAppPath & "\Informes.ini"), cAppPath, App.Path)
    
    If Not PathExists(Cfg.Data) Then
        MsgBox "El archivo de Datos no existe o no se puede tener acceso a él." & vbNewLine & vbNewLine & _
        "Se creará uno nuevo en la ruta del programa. Modifique los valores según las instrucciones del programa.", vbCritical, "No se encuentra el archivo"
        Call CreateDataFile
    End If
        
    ' Si no existen simultaneamente los archivos de Informes, Pedidos y Enlace, inicializarlos
    If Not PathExists(Cfg.Reports) And Not PathExists(Cfg.Tasks) And Not PathExists(Cfg.Link) Then
        MsgBox "Los archivos de Pedidos, Informes y Enlace no existen, por lo que se inicializará el Sistema de Pedidos." & vbNewLine & vbNewLine & _
        "Se crearán los archivos nuevos en la ruta del programa, para cambiarlos de ubicación u nombre, modifique el archivo " & Cfg.Paths & ".", vbInformation, "Modo de inicialización"
        Call InitializeReportsTasksAndLink
        fAbortBackup = True
    Else
    ' Si alguno no existe pero el resto si, notificar
        If Not PathExists(Cfg.Reports) Then bTasksReportsOrLinkFailed = True: sWhichFailed = sWhichFailed & Cfg.Reports & vbNewLine
        If Not PathExists(Cfg.Tasks) Then bTasksReportsOrLinkFailed = True: sWhichFailed = sWhichFailed & Cfg.Tasks & vbNewLine
        If Not PathExists(Cfg.Link) Then bTasksReportsOrLinkFailed = True: sWhichFailed = sWhichFailed & Cfg.Link & vbNewLine
    
        If bTasksReportsOrLinkFailed Then
            MsgBox "No se puede obtener acceso a el/los archivo(s):" & vbNewLine & sWhichFailed & vbNewLine & _
            "Si intentó reinicializar el programa para comenzar de cero, borre simultáneamente estos tres archivos. De lo contrario, restaure una copia de seguridad de los mismos." & _
            vbNewLine & vbNewLine & "El programa se cerrará.", vbExclamation, "Ocurrió un error al leer los archivos de datos"
            Destroy (True)
        End If
    End If

    Dim i As Long
    Dim s As String
    
    ' Si deciden usar backups y no fallo ningun archivo
    MaxBackups = Val(ReadIni(Cfg.Data, "BACKUPS", "CantidadDeCopias", 3))
    fUseBuiltInBackup = IIf(ReadIni(Cfg.Data, "BACKUPS", "Usar", 0) = 0, False, True)
    BackupsPath = FormatPath(Replace(ReadIni(Cfg.Data, "BACKUPS", "Ruta", cAppPath & "\Backups"), cAppPath, App.Path))
    
    If fUseBuiltInBackup Then
        If Not PathExists(BackupsPath) Then MkDir (BackupsPath)
        
        If Not fAbortBackup Then
            For i = MaxBackups To 1 Step -1
                If PathExists(BackupsPath & "Informes" & i & ".ini.bak") Then Name BackupsPath & "Informes" & i & ".ini.bak" As BackupsPath & "Informes" & i + 1 & ".ini.bak"
                If PathExists(BackupsPath & "Pedidos" & i & ".ini.bak") Then Name BackupsPath & "Pedidos" & i & ".ini.bak" As BackupsPath & "Pedidos" & i + 1 & ".ini.bak"
                If PathExists(BackupsPath & "Enlace" & i & ".ini.bak") Then Name BackupsPath & "Enlace" & i & ".ini.bak" As BackupsPath & "Enlace" & i + 1 & ".ini.bak"
            Next
             
           If PathExists(BackupsPath & "Informes" & MaxBackups + 1 & ".ini.bak") Then Kill BackupsPath & "Informes" & MaxBackups + 1 & ".ini.bak"
           If PathExists(BackupsPath & "Pedidos" & MaxBackups + 1 & ".ini.bak") Then Kill BackupsPath & "Pedidos" & MaxBackups + 1 & ".ini.bak"
           If PathExists(BackupsPath & "Enlace" & MaxBackups + 1 & ".ini.bak") Then Kill BackupsPath & "Enlace" & MaxBackups + 1 & ".ini.bak"
        
           FileCopy Cfg.Reports, BackupsPath & "Informes1.ini.bak"
           FileCopy Cfg.Tasks, BackupsPath & "Pedidos1.ini.bak"
           FileCopy Cfg.Link, BackupsPath & "Enlace1.ini.bak"
        End If
    End If
    
    ' Debug
    If PathExists(App.Path & "\Debug.rnm") Then
        fDebug = True
    End If
    
    ' Dimensiones de frmMain
    frmMainW = CLng(ReadIni(Cfg.Data, "MAIN", "frmW", "13500"))
    frmMainH = CLng(ReadIni(Cfg.Data, "MAIN", "frmH", "7200"))
    WHThreshold = CSng(ReadIni(Cfg.Data, "MAIN", "WHThreshold", "0,8"))

    ' Fuentes de la interfaz
    For i = 0 To 4
        sFontName(i) = ReadIni(Cfg.Data, "FONT", "Fuente" & i, "Courier New")
        sFontSize(i) = Val(ReadIni(Cfg.Data, "FONT", "Grado" & i, "11"))
    Next i
        
        fGlobalFont = ReadIni(Cfg.Data, "FONT", "FuenteUnica", "0")
        gSepChar = ReadIni(Cfg.Data, "FONT", "gSepChar", "2")
    
    ' Otros datos
    SubjectLen = Val(ReadIni(Cfg.Data, "EXTRA", "SubjectLen", 80))
    SubjectFormat = ReadIni(Cfg.Data, "EXTRA", "SubjectFormat", "%P - %D")
    fClipboardCopyNotify = IIf(ReadIni(Cfg.Data, "EXTRA", "ClpbrdNotify", "1") = "0", False, True)
    
    i = -1
    Do
        i = i + 1
        s = ReadIni(Cfg.Data, "TECNICOS", i, "")
    
        If s = "" Then i = i - 1: Exit Do
        If i > 999 Then i = 0: Exit Do
    Loop
     
    TechnicianAmount = i
    
    Load frmMain
    Exit Sub
    
SomethingWentWrong:
MsgBox "Ocurrió un error al ejecutar las tareas iniciales del Sistema de Pedidos (Error " & Err.Number & "). La descripción del error es la siguiente: «" & _
        Err.Description & "»." & vbNewLine & vbNewLine & _
        "Por favor, verifique que los archivos esenciales del programa se encuentran en las rutas especificadas en el archivo Rutas.ini y que tiene acceso a ellos." & vbNewLine & vbNewLine & _
        "Este programa se cerrará.", vbCritical, "Error " & Err.Number & " al iniciar el Sistema de Pedidos"
        
    End
End Sub

Function Destroy(Optional ByVal Force As Boolean = False) As Boolean
' Se llama desde cada instancia en la que se quiere cerrar el programa

    If Not Force Then
        If MsgBox("¿Está seguro de que desea salir?", vbQuestion + vbYesNo, "Salir") = vbNo Then
            Destroy = False
            Exit Function
        End If
       
        WriteIni Cfg.Data, "FLAGS", "Mostrar", MainFilter
        WriteIni Cfg.Data, "FLAGS", "DestacarPedidosConInformes", CLng(fHlTasksWithReports)
        WriteIni Cfg.Data, "FLAGS", "MostrarPrioridad", CLng(fShowPriority)
        WriteIni Cfg.Data, "TECNICOS", "UltimoLogueado", CurrentTechnicianIndex
        WriteIni Cfg.Data, "MAIN", "FrmW", frmMain.NormalizeWidth
        WriteIni Cfg.Data, "MAIN", "FrmH", frmMain.NormalizeHeight
    End If

   Destroy = True
    
    Unload frmNewReport
    Unload frmNewTask
    Unload frmPickTechnician
    Unload frmReportViewer
    Unload frmSearch
    Unload frmTaskViewer
    Unload frmSplash
    Unload frmWelcomeSplash
    Unload frmOptions
    Unload frmDebug
    Unload frmMain

    End
End Function

Sub CreatePathFile()
Dim sPath As String
On Error Resume Next

sPath = App.Path & "\Rutas.ini"

    WriteIni sPath, "MAIN", "Datos", "&f\Datos.ini"
    WriteIni sPath, "MAIN", "Enlace", "&f\Enlace.ini"
    WriteIni sPath, "MAIN", "Informes", "&f\Informes.ini"
    WriteIni sPath, "MAIN", "Pedidos", "&f\Pedidos.ini"
End Sub

Sub CreateDataFile()
Dim sPath As String
On Error Resume Next

sPath = App.Path & "\Datos.ini"

    WriteIni sPath, "SUCURSALES", "0", "Central"
    WriteIni sPath, "TECNICOS", "0", "FGF"
    WriteIni sPath, "TECNICOS", "UltimoLogueado", "0"
    WriteIni sPath, "ASOCIAR_TECNICO", "ffigun", "0"
    WriteIni sPath, "FLAGS", "Mostrar", "1"
    WriteIni sPath, "FLAGS", "MostrarPrioridad", "0"
    WriteIni sPath, "FLAGS", "DestacarPedidosConInformes", "0"
    WriteIni sPath, "FGWIKI", "Mostrar", "0"
    WriteIni sPath, "FGWIKI", "RutaExe", ""
    WriteIni sPath, "BACKUPS", "Usar", "1"
    WriteIni sPath, "BACKUPS", "CantidadDeCopias", "4"
    WriteIni sPath, "BACKUPS", "Ruta", "&f\Backups"
    WriteIni sPath, "FONT", "FuenteUnica", "1"
    WriteIni sPath, "FONT", "Fuente0", "Calibri"
    WriteIni sPath, "FONT", "Fuente1", "Lucida Sans Typewriter"
    WriteIni sPath, "FONT", "Fuente2", "Lucida Sans Typewriter"
    WriteIni sPath, "FONT", "Fuente3", "Calibri"
    WriteIni sPath, "FONT", "Fuente4", "Consolas"
    WriteIni sPath, "FONT", "Grado0", "12"
    WriteIni sPath, "FONT", "Grado1", "11"
    WriteIni sPath, "FONT", "Grado2", "11"
    WriteIni sPath, "FONT", "Grado3", "12"
    WriteIni sPath, "FONT", "Grado4", "11"
    WriteIni sPath, "FONT", "gSepChar", "2"
    WriteIni sPath, "SEARCH", "frmW", "9600"
    WriteIni sPath, "SEARCH", "frmH", "5820"
    WriteIni sPath, "MAIN", "frmW", "13500"
    WriteIni sPath, "MAIN", "frmH", "7200"
    WriteIni sPath, "MAIN", "WHThreshold", "0,8"
    WriteIni sPath, "EXTRA", "SubjectLen", "80"
    WriteIni sPath, "EXTRA", "ClpbrdNotify", "0"
    WriteIni sPath, "EXTRA", "SubjectFormat", "%N - %D"

End Sub

Sub InitializeReportsTasksAndLink()
    Dim ff As Integer
    
    ff = FreeFile
    Open Cfg.Link For Output As ff: Close ff
    
    ff = FreeFile
    Open Cfg.Reports For Output As ff: Close ff
    
    ff = FreeFile
    Open Cfg.Tasks For Output As ff: Close ff
End Sub

Function GetCurrentUserName(Optional ByVal sMinus As Boolean = False) As String
Dim lpBuff   As String * 25
Dim ret      As Long
Dim UserName As String

' Obtener el nombre de usuario y quitar los espacios
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
    UserName = IIf(sMinus, LCase(UserName), UserName)
    
    GetCurrentUserName = UserName
End Function

Function GetAssociatedTechnician(ByVal WindowsUserName As String) As Integer
' Devuelve "" si no tiene nada asociado
    GetAssociatedTechnician = Val(ReadIni(Cfg.Data, "ASOCIAR_TECNICO", WindowsUserName, -1))
End Function

Function ExportSpecificTask(ByVal sPath As String, ByVal Number As Long) As Boolean
On Error GoTo SadPanda

Dim ff          As Integer: ff = FreeFile   ' Número de archivo libre de Windows
Dim i           As Long                     ' Bucle
Dim j           As Long                     ' Bucle
Dim fn          As String                   ' Nombre de archivo
Dim tt          As TaskStructure            ' Pedido temporal
Dim tr          As ReportStructure          ' Informe temporal
Dim bf()        As String                   ' Buffer

Dim aReports()  As Long                     ' Contenedor de los informes del pedido
Dim cReports    As Long                     ' Cantidad de informes
    
    If Not PathExists(FormatPath(sPath)) Then GoTo SadPanda
    If Not TaskExists(Number) Then GoTo SadPanda

sPath = FormatPath(sPath)
fn = sPath & "Pedido " & FormatTaskNumber(Number) & ".txt"

Screen.MousePointer = vbHourglass

Open fn For Output As #ff
    tt = ReadTask(Number)
    aReports = GetTaskReports(CurrentTask)
    cReports = UBound(aReports)
    
    With tt
        Print #ff, "*" & String(78, "-") & "*"
        Print #ff, "| " & PadStr("Pedido: " & FormatTaskNumber(.Number), 35) & PadStr("Fecha: " & FormatDate2(.Date), 20) & PadStr("Hora: " & FormatTime(.Time), 21) & " |"
        Print #ff, "| " & PadStr("Contacto: " & .Contact, 35) & PadStr("Fecha Ok: " & FormatDate2(.OkDate), 20) & PadStr("Informe Ok: " & FormatReportNumber(.OkReport), 21) & " |"
        Print #ff, "| " & PadStr("Sucursal: " & .BranchOffice, 35) & PadStr("Técnico: " & .Technician, 20) & PadStr("Estado: " & .Status, 21) & " |"
        Print #ff, "| " & Space(55) & PadStr("Prioridad: " & .Priority, 21) & " |"
        Print #ff, "| " & Space(76) & " |"
        
        Print #ff, "| " & PadStr("Detalle:", 76) & " |"
        bf = CustomWordWrap(.Detail, 80)
            For i = LBound(bf) To UBound(bf)
                Print #ff, bf(i)
            Next i
        
        Print #ff, "| " & Space(76) & " |"
        
        Print #ff, "| " & PadStr("Observaciones:", 76) & " |"
        bf = CustomWordWrap(.Observations, 80)
        Print #ff, bf(0)
        
        Print #ff, "| " & Space(76) & " |"
        Print #ff, "*" & String(78, "-") & "*"
    End With
    
    For j = 1 To UBound(aReports)
        tr = ReadReport(aReports(j))
        
        With tr
            Print #ff, "| " & PadStr("Informe: " & FormatReportNumber(.Number), 35) & PadStr("Fecha: " & FormatDate2(.Date), 25) & PadStr("Hora: " & FormatTime(.Time), 16) & " |"
            Print #ff, "| " & PadStr("Contacto: " & .Contact, 35) & PadStr("Sucursal: " & .BranchOffice, 25) & PadStr("Técnico: " & .Technician, 16) & " |"
            Print #ff, "| " & Space(76) & " |"
            Print #ff, "| " & PadStr("Detalle:", 76) & " |"
            
                bf = CustomWordWrap(.Detail, 80)
                For i = LBound(bf) To UBound(bf)
                    Print #ff, bf(i)
                Next i
            
            Print #ff, "| " & Space(76) & " |"
            Print #ff, "*" & String(78, "-") & "*"
        End With
    Next j
Close #ff

Screen.MousePointer = vbDefault

' Variable global
strBufferPath = fn
ExportSpecificTask = True

Exit Function

SadPanda:
    Screen.MousePointer = vbNormal
    ExportSpecificTask = False
End Function

Function ExportTasks(ByVal sPath As String) As Boolean
On Error GoTo SadPanda

Dim ff          As Integer: ff = FreeFile   ' Número de archivo libre de Windows
Dim i           As Long                     ' Bucle
Dim fn          As String                   ' Nombre de archivo
Dim tt          As TaskStructure            ' Pedido temporal

Dim j           As Integer
Dim tBuffer(3)  As String

Const Detail = 0
Const Branch = 1
Const Contac = 2
Const Observ = 3

    If Not PathExists(FormatPath(sPath)) Then GoTo SadPanda

sPath = FormatPath(sPath)
fn = sPath & "Pedidos_Exportados_" & PlainDate(Date) & ".csv"

Screen.MousePointer = vbHourglass

Open fn For Output As #ff
    Write #ff, "#", "Contacto", "Sucursal", "Fecha", "Detalle", "Estado", "Técnico", "Prioridad", "Cant. Informes", "Fecha OK", "Informe OK", "Observaciones"
    For i = 0 To LastTask
        If TaskExists(i) Then
        tt = ReadTask(i)
        
        With tt
            tBuffer(Detail) = Replace(.Detail, vbNewLine, " ")
            tBuffer(Branch) = Replace(.BranchOffice, vbNewLine, " ")
            tBuffer(Contac) = Replace(.Contact, vbNewLine, " ")
            tBuffer(Observ) = Replace(.Observations, vbNewLine, " ")
                    
            For j = 0 To 3
                tBuffer(j) = Replace(tBuffer(j), Chr(34), "''")
                tBuffer(j) = Replace(tBuffer(j), ";", ",")
            Next j
            
            Write #ff, FormatTaskNumber(.Number), _
                       tBuffer(Contac), _
                       tBuffer(Branch), _
                       FormatDate2(.Date), _
                       tBuffer(Detail), _
                       .Status, _
                       .Technician, _
                       Str(.Priority), _
                       .ReportAmount, _
                       FormatDate2(.OkDate), _
                       IIf(.OkReport = "-1", "", FormatReportNumber(.OkReport)), _
                       tBuffer(Observ)

            For j = 0 To 3
                tBuffer(j) = ""
            Next j
        
            End With
        End If
    Next
Close #ff

Screen.MousePointer = vbDefault

' Variable global
strBufferPath = fn

ExportTasks = True
Erase tBuffer

Exit Function

SadPanda:
    Screen.MousePointer = vbNormal
    ExportTasks = False
    Erase tBuffer
End Function

Function ExportReports(ByVal sPath As String) As Boolean
On Error GoTo SadPanda

Dim ff          As Integer: ff = FreeFile   ' Número de archivo libre de Windows
Dim i           As Long                     ' Bucle
Dim fn          As String                   ' Nombre de archivo
Dim tr          As ReportStructure          ' Informe temporal

Dim j           As Integer
Dim tBuffer(2)  As String

Const Detail = 0
Const Branch = 1
Const Contac = 2

    If Not PathExists(FormatPath(sPath)) Then GoTo SadPanda
    
sPath = FormatPath(sPath)
fn = sPath & "Informes_Exportados_" & PlainDate(Date) & ".csv"

Screen.MousePointer = vbHourglass

Open fn For Output As #ff
    Write #ff, "#", "Pedido de Referencia", "Contacto", "Sucursal", "Fecha", "Detalle", "Técnico"
    
    For i = 0 To LastTask
        If ReportExists(i) Then
            tr = ReadReport(i)
            With tr
                tBuffer(Detail) = Replace(.Detail, vbNewLine, " ")
                tBuffer(Branch) = Replace(.BranchOffice, vbNewLine, " ")
                tBuffer(Contac) = Replace(.Contact, vbNewLine, " ")
                
                For j = 0 To 2
                   tBuffer(j) = Replace(tBuffer(j), Chr(34), "''")
                   tBuffer(j) = Replace(tBuffer(j), ";", ",")
                Next j
                            
                Write #ff, FormatReportNumber(.Number), _
                           FormatTaskNumber(.ReferenceTask), _
                           tBuffer(Contac), _
                           tBuffer(Branch), _
                           FormatDate2(.Date), _
                           tBuffer(Detail), _
                           .Technician

                For j = 0 To 2
                    tBuffer(j) = ""
                Next j
            End With
        End If
    Next
Close #ff

Screen.MousePointer = vbDefault

' Variable global
strBufferPath = fn

ExportReports = True
Erase tBuffer

Exit Function

SadPanda:
    Screen.MousePointer = vbNormal
    ExportReports = False
    Erase tBuffer
End Function

Function PathExists(ByVal sPath As String) As Boolean
On Error GoTo DoesNotExist
    PathExists = IIf(Dir$(sPath, vbDirectory) <> vbNullString, True, False)
    Exit Function

DoesNotExist: PathExists = False
End Function

Function FormatPath(ByVal sPath As String) As String
' Agrega \ al final según corresponda
    FormatPath = IIf(Right(sPath, 1) = "\", sPath, sPath & "\")
End Function

Function IsTimeValid(ByVal TimeAs_hhmm As String) As Boolean
On Error GoTo IsNotValid
    
    IsTimeValid = IsDate(CDate(Left$(TimeAs_hhmm, 2) & ":" & Right$(TimeAs_hhmm, 2)))
    Exit Function

IsNotValid:
IsTimeValid = False
End Function

Function IsDateValid(ByVal DateAs_ddmmyy As String) As Boolean
On Error GoTo IsNotValid
Dim sDateBuffer As Date

' mm/dd/yy
    sDateBuffer = Mid$(DateAs_ddmmyy, 3, 2) & "/" & Mid$(DateAs_ddmmyy, 1, 2) & "/" & Mid$(DateAs_ddmmyy, 5, 2)
    
    If IsDate(sDateBuffer) Then
        IsDateValid = True
    Else
        IsDateValid = False
    End If
       
Exit Function
    
IsNotValid:
    IsDateValid = False
End Function

Function ReadTask(ByVal Number As Long) As TaskStructure
Dim tsBuffer            As TaskStructure
Dim TaskDataHolder()    As String
    
Dim a   As String: a = Cfg.Tasks
Dim b   As String: b = FormatTaskNumber(Number)
Dim c   As String: c = Cfg.Link

' Levantar datos del key DATA del pedido y parsearlo
    ReDim TaskDataHolder(UBoundTaskData)
    TaskDataHolder = Split(ReadIni(a, b, cData, ""), cSepChar)
        
    With tsBuffer
        .Number = FormatTaskNumber(Number)
       
        .Contact = TaskDataHolder(TaskDataIndex.Contact)
        .Detail = Replace(TaskDataHolder(TaskDataIndex.Detail), cNewLine, vbNewLine)
        .Status = TaskDataHolder(TaskDataIndex.Status)
        .OkDate = Trim(TaskDataHolder(TaskDataIndex.OkDate))
        .Date = TaskDataHolder(TaskDataIndex.InputDate)
        .Time = TaskDataHolder(TaskDataIndex.InputTime)
        .Observations = TaskDataHolder(TaskDataIndex.Observations)
        .Priority = TaskDataHolder(TaskDataIndex.Priority)
        .BranchOffice = TaskDataHolder(TaskDataIndex.BranchOffice)
        .Technician = TaskDataHolder(TaskDataIndex.Technician)
        .OkReport = TaskDataHolder(TaskDataIndex.OkReport)

        .ReportAmount = ReadIni(c, b, cReportsQty, 0)
    End With
    
    ReadTask = tsBuffer
End Function

Function WriteTask(ByRef tsTask As TaskStructure) As Boolean
On Error GoTo OhNo
Dim TaskDataBuffer As String
    
Dim a As String: a = Cfg.Tasks
Dim b As String: b = FormatTaskNumber(tsTask.Number)

    With tsTask
    .Detail = Replace(.Detail, vbNewLine, cNewLine)
        
        TaskDataBuffer = .Contact & cSepChar & _
                         .Detail & cSepChar & _
                         .Status & cSepChar & _
                         .OkDate & cSepChar & _
                         .Date & cSepChar & _
                         .Time & cSepChar & _
                         .Observations & cSepChar & _
                         .Priority & cSepChar & _
                         .BranchOffice & cSepChar & _
                         .Technician & cSepChar & _
                         .OkReport
    End With
    
    WriteIni a, b, cData, TaskDataBuffer

    WriteTask = True
    
    Exit Function

OhNo:
    WriteTask = False
End Function

Function ReadReport(ByVal Number As Long) As ReportStructure
Dim rsBuffer            As ReportStructure
Dim ReportDataHolder()  As String

Dim a As String: a = Cfg.Reports
Dim b As String: b = FormatReportNumber(Number)
Dim c As String: c = "<N/A>"
    
    ReDim ReportDataHolder(UBoundReportData)
    ReportDataHolder = Split(ReadIni(a, b, cData, "-1"), cSepChar)
    
    With rsBuffer
        .Number = FormatReportNumber(Number)
       
        .Contact = ReportDataHolder(ReportDataIndex.Contact)
        .Detail = Replace(ReportDataHolder(ReportDataIndex.Detail), cNewLine, vbNewLine)
        .Date = ReportDataHolder(ReportDataIndex.InputDate)
        .Time = ReportDataHolder(ReportDataIndex.InputTime)
        .BranchOffice = ReportDataHolder(ReportDataIndex.BranchOffice)
        .Technician = ReportDataHolder(ReportDataIndex.Technician)
        .ReferenceTask = ReportDataHolder(ReportDataIndex.ReferenceTask)
    End With
    
    ReadReport = rsBuffer
End Function

Function WriteReport(ByRef rsReport As ReportStructure) As Boolean
On Error GoTo OhNo
    
Dim ReportDataBuffer As String

Dim a As String: a = Cfg.Reports
Dim b As String: b = FormatReportNumber(rsReport.Number)
    
    With rsReport
    .Detail = Replace(.Detail, vbNewLine, cNewLine)
    
        ReportDataBuffer = .Contact & cSepChar & _
                           .Detail & cSepChar & _
                           .Date & cSepChar & _
                           .Time & cSepChar & _
                           .BranchOffice & cSepChar & _
                           .Technician & cSepChar & _
                           .ReferenceTask
                           
        Link .ReferenceTask, .Number
    End With
    
    WriteIni a, b, cData, ReportDataBuffer
    WriteReport = True

Exit Function

OhNo:
    WriteReport = False
End Function

Function GetTaskData(ByVal Number As Long) As String
    GetTaskData = ReadIni(Cfg.Tasks, FormatTaskNumber(Number), cData, "")
End Function

Function GetReportData(ByVal Number As Long) As String
    GetReportData = ReadIni(Cfg.Reports, FormatReportNumber(Number), cData, "")
End Function

Function GetTime(ByVal Number As Long, ByVal OfTask As Boolean) As Date
Dim sBuffer() As String

    If OfTask Then
        sBuffer = Split(GetTaskData(Number), cSepChar)
        GetTime = Format(sBuffer(TaskDataIndex.InputTime), "HHmm")
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetTime = Format(sBuffer(ReportDataIndex.InputTime), "HHmm")
    End If
End Function

Function GetStatus(ByVal Number As Long) As String
Dim sBuffer() As String
    sBuffer = Split(GetTaskData(Number), cSepChar)
    GetStatus = sBuffer(TaskDataIndex.Status)
End Function

Function GetReportAmount(ByVal Number As Long) As Long
   GetReportAmount = ReadIni(Cfg.Link, FormatTaskNumber(Number), cReportsQty, 0)
End Function

Function GetOkReport(ByVal Number As Long) As Long
Dim sBuffer() As String
    sBuffer = Split(GetTaskData(Number), cSepChar)
    GetOkReport = sBuffer(TaskDataIndex.OkReport)
End Function

Function GetOkDate(ByVal Number As Long) As String
Dim sBuffer() As String
    sBuffer = Split(GetTaskData(Number), cSepChar)
    GetOkDate = sBuffer(TaskDataIndex.OkDate)
End Function

Function GetDetail(ByVal Number As Long, ByVal OfTask As Boolean) As String
Dim sBuffer() As String
    If OfTask Then
        sBuffer = Split(GetTaskData(Number), cSepChar)
        GetDetail = sBuffer(TaskDataIndex.Detail)
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetDetail = sBuffer(ReportDataIndex.Detail)
    End If
End Function

Function GetReferenceTask(ByVal Number As Long) As Long
Dim sBuffer() As String
    sBuffer = Split(GetReportData(Number), cSepChar)
    GetReferenceTask = sBuffer(ReportDataIndex.ReferenceTask)
End Function

Function GetContact(ByVal Number As Long, ByVal OfTask As Boolean, Optional Normalizar As Boolean = False) As String
Dim sBuffer() As String
    If OfTask Then
        sBuffer = Split(GetTaskData(Number), cSepChar)
        GetContact = sBuffer(TaskDataIndex.Contact)
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetContact = sBuffer(ReportDataIndex.Contact)
    End If
End Function

Function GetBranchOffice(ByVal Number As Long, ByVal OfTask As Boolean) As String
Dim sBuffer() As String
    If OfTask Then
        sBuffer = Split(GetTaskData(Number), cSepChar)
        GetBranchOffice = sBuffer(TaskDataIndex.BranchOffice)
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetBranchOffice = sBuffer(ReportDataIndex.BranchOffice)
    End If
End Function

Function GetPriority(ByVal Number As Long) As Integer
Dim sBuffer() As String
    sBuffer = Split(GetTaskData(Number), cSepChar)
    GetPriority = sBuffer(TaskDataIndex.Priority)
End Function

Function GetTechnician(ByVal Number As Long, ByVal OfTask As Boolean) As String
Dim sBuffer() As String
    If OfTask Then
        sBuffer = Split(GetTaskData(Number), cSepChar)
        GetTechnician = sBuffer(TaskDataIndex.Technician)
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetTechnician = sBuffer(ReportDataIndex.Technician)
    End If
End Function

Function GetTaskReports(ByVal Number As Long) As Long()
Dim i As Long            ' Bucle
Dim cReports As Long     ' Cantidad de informes
Dim Buffer() As Long     ' Buffer
    
    cReports = GetReportAmount(FormatTaskNumber(Number))

    ReDim Buffer(cReports)
    
    For i = LBound(Buffer) To UBound(Buffer)
    ' Si GetRelativeReport es 0, devuelve -1
        Buffer(i) = GetRelativeReport(Number, i)
    Next i
    
    GetTaskReports = Buffer
End Function

Function GetRelativeReport(ByVal Number As Long, Index As Long) As Long
' Devuelve la posición del informe respecto a los informes de un pedido
    GetRelativeReport = CLng(ReadIni(Cfg.Link, FormatTaskNumber(Number), cReport & Index, "-1"))
End Function

Function GetObservations(ByVal Number As Long) As String
Dim sBuffer() As String
    sBuffer = Split(GetTaskData(Number), cSepChar)
    GetObservations = sBuffer(TaskDataIndex.Observations)
End Function

Function GetDate(ByVal Number As Long, ByVal OfTask As Boolean, Optional ByVal OkDate As Boolean = False) As String
' Devuelve la fecha de inicio o fin de un pedido o informe como STRING

' Si hay que dar la fecha de cierre pero el pedido no esta OK, salir
    If OkDate = True Then
        If UCase(GetStatus(Number)) <> strOK Then
            Exit Function
        End If
    End If

' Los informes no tienen fecha de finalizacion
    If OkDate = True Then
        If OfTask = False Then
            Exit Function
        End If
    End If
   
   
Dim sBuffer() As String
    If OfTask Then
    sBuffer = Split(GetTaskData(Number), cSepChar)
        If OkDate Then
            GetDate = sBuffer(TaskDataIndex.OkDate)
        Else
            GetDate = sBuffer(TaskDataIndex.InputDate)
        End If
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        GetDate = sBuffer(ReportDataIndex.InputDate)
    End If
End Function

Function GetDate2(ByVal Number As Long, ByVal OfTask As Boolean, Optional ByVal OkDate As Boolean = False) As Date
' Devuelve la fecha de inicio o fin de un pedido o informe como DATE

' Si hay que dar la fecha de cierre pero el pedido no esta OK, salir
    If OkDate = True Then
        If UCase(GetStatus(Number)) <> strOK Then
            Exit Function
        End If
    End If

' Los informes no tienen fecha de finalizacion
    If OkDate = True Then
        If OfTask = False Then
            Exit Function
        End If
    End If

Dim sBuffer()   As String   ' Parse
Dim sTemp       As String   ' Holder

    If OfTask Then
    sBuffer = Split(GetTaskData(Number), cSepChar)
        If OkDate Then
            sTemp = sBuffer(TaskDataIndex.OkDate)
        Else
            sTemp = sBuffer(TaskDataIndex.InputDate)
        End If
    Else
        sBuffer = Split(GetReportData(Number), cSepChar)
        sTemp = sBuffer(ReportDataIndex.InputDate)
    End If


   If sTemp = "-1" Then Exit Function
   If Trim(sTemp) = "" Then Exit Function
   If Len(Trim(sTemp)) <> 6 Then Exit Function
    
   GetDate2 = Format(sTemp, "@@/@@/@@")
End Function

Function FormatDate(ByVal DateAs_ddmmyy As String) As Date
' Devuelve un DATE a lo que ingresan los usuarios en las cajas de texto de fecha
If Not IsNumeric(DateAs_ddmmyy) Then Exit Function

    FormatDate = Format(DateAs_ddmmyy, "@@/@@/@@")
End Function

Function FormatDate2(ByVal pDate As String) As String
' Devuelve dd/mm/yy como STRING en base a ddmmyy
If Not IsNumeric(pDate) Then Exit Function

    FormatDate2 = Mid$(pDate, 1, 2) & _
            "/" & Mid$(pDate, 3, 2) & _
            "/" & Mid$(pDate, 5, 2)
End Function

Function FormatTime(ByVal TimeAs_hhmm As String) As String
On Error GoTo Simple
    
    FormatTime = Format(CDate(Left$(TimeAs_hhmm, 2) & ":" & Right$(TimeAs_hhmm, 2)), "HH:mm")
    Exit Function

Simple:
    FormatTime = Left$(TimeAs_hhmm, 2) & ":" & Right$(TimeAs_hhmm, 2)
End Function

Sub Link(ByVal TaskNumber As Long, ByVal ReportNumber As Long)
Dim a As String: a = Cfg.Link
Dim b As String: b = FormatTaskNumber(TaskNumber)
Dim c As String: c = FormatReportNumber(ReportNumber)

Dim i As Long
Dim j As Long

' Leer la cantidad de informes que tiene el pedido. Si no tiene ninguna, establece 0
    i = Val(ReadIni(a, b, cReportsQty, "0"))

' Si ya está el informe entre los informes, no ejecutar la funcion
    If Not i <= 0 Then
        For j = 1 To i
            If CLng(ReadIni(a, b, cReport & j, "0")) = ReportNumber Then
                Exit Sub
            End If
        Next
    End If

' Añadir un valor llamado InformeX donde X es el siguiente
    If Trim(ReadIni(a, b, cReport & i + 1, "")) <> "" Then
        MsgBox "El programa está tratando de escribir un registro en un campo ya existente. Esto sobreescribiría los datos." & vbNewLine & vbNewLine & "Por favor, compruebe la integridad del archivo de Enlaces. El informe no se guardará.", vbCritical, "Error de integridad"
        Exit Sub
    End If
    
' Si pasa la comprobacion, se escribe el numero de informe correspondiente
    WriteIni a, b, cReportsQty, i + 1
    WriteIni a, b, cReport & i + 1, c
End Sub

Function TaskExists(ByVal Number As Long) As Boolean
    If Trim(GetTaskData(Number)) <> "" Then
        TaskExists = True
    Else
        TaskExists = False
    End If
End Function

Function ReportExists(ByVal Number As Long) As Boolean
    If Trim(GetReportData(Number)) <> "" Then
        ReportExists = True
    Else
        ReportExists = False
    End If
End Function

Function CloseTask(ByVal Number As Long, ByVal ReportNumber As Long, ByVal OkDate As String) As Boolean
' Considerar que los parametros indicados son correctos y existen
On Error GoTo OhNo

If GetStatus(Number) = strOK Then Exit Function

Dim TaskBuffer As TaskStructure
    TaskBuffer = ReadTask(Number)
    
With TaskBuffer
    .OkReport = FormatReportNumber(ReportNumber)
    .OkDate = OkDate
    .Status = strOK
End With

    WriteTask TaskBuffer
    CloseTask = True
Exit Function
       
OhNo:
    MsgBox "No se pudo cerrar el pedido " & FormatTaskNumber(Number) & ". El error devuelto fue: (" & Err.Number & ") " & Err.Description, vbExclamation, "Error " & Err.Number
    CloseTask = False
End Function

Function OpenClosedTask(ByVal Number As Long) As Boolean
' Considerar que los parametros indicados son correctos y existen
On Error GoTo OhNo

If GetStatus(Number) <> strOK Then Exit Function

Dim TaskBuffer As TaskStructure
    TaskBuffer = ReadTask(Number)
    
With TaskBuffer
    .OkReport = "-1"
    .OkDate = ""
    .Status = strPEND
End With

    WriteTask TaskBuffer
    OpenClosedTask = True
Exit Function

OhNo:
    MsgBox "No se pudo abrir el pedido " & FormatNumber(Number) & ". El error devuelto fue: (" & Err.Number & ") " & Err.Description, vbExclamation, "Error " & Err.Number
    OpenClosedTask = False
End Function

Function PadStr(ByVal strSource As String, ByVal lPadLen As Long, Optional ByVal PadChar As String = " ") As String
' Llena los espacios faltantes con caracteres PadChar
    PadStr = String(lPadLen, Left(PadChar, 1))
    Mid(PadStr, 1, Len(strSource)) = strSource
End Function

Function PlainDate(ByVal pDate As Date)
' ddmmyy
    PlainDate = Format(pDate, "ddmmyy")
End Function

Function PlainTime(ByVal pHour As Date)
' hhmm
    PlainTime = Format(pHour, "hhmm")
End Function

 Function FormatTaskNumber(ByVal Number As Long) As String
    If Number = -1 Then
        FormatTaskNumber = ""
    Else
        FormatTaskNumber = Format(Str(Number), "00000")
    End If
End Function

Function FormatReportNumber(ByVal Number As Long) As String
    If Number = -1 Then
        FormatReportNumber = ""
    Else
        FormatReportNumber = Format(Str(Number), "00000")
    End If
End Function

Function FreeTask() As Long
    FreeTask = LastTask + 1
End Function

Function LastTask() As Long
' -2: Error | -1: Sin pedidos cargados
    LastTask = Val(ReadIni(Cfg.Link, "Last", cTask, -2))
End Function

Function LastTaskFull() As Long
Dim fEOF As Boolean: fEOF = False
Dim i As Long: i = -1

    If Not TaskExists(0) Then
        fEOF = True
        LastTaskFull = -1
        Exit Function
    End If

    Do Until fEOF = True Or i > 2147483646
      i = i + 1
        If Not TaskExists(i) Then
            i = i - 1
            fEOF = True
        End If
    Loop
    LastTaskFull = i
End Function

Sub AddOneToLastTask()
Dim lt As Long
    lt = LastTask
    
    If lt = -2 Then
        MsgBox "Ocurrió un error al actualizar el índice del último pedido." & vbNewLine & vbNewLine & _
               "Se recomienda no utilizar el programa para prevenir la sobreescritura de datos de otros índices.", vbCritical, "Error al actualizar el índice"
        Exit Sub
    Else
        WriteIni Cfg.Link, "Last", cTask, lt + 1
    End If
End Sub

Function FreeReport() As Long
    FreeReport = LastReport + 1
End Function

Function LastReport() As Long
' -2: Error | -1: Sin informes cargados
    LastReport = Val(ReadIni(Cfg.Link, "Last", cReport, -2))
End Function

Function LastReportFull() As Long
Dim fEOF As Boolean: fEOF = False
Dim i As Long: i = -1

    If Not ReportExists(0) Then
        fEOF = True
        LastReportFull = -1
        Exit Function
    End If

    Do Until fEOF = True Or i > 2147483646
      i = i + 1
        If Not ReportExists(i) Then
            i = i - 1
            fEOF = True
        End If
    Loop
    LastReportFull = i
End Function

Sub AddOneToLastReport()
Dim lr As Long
    lr = LastReport
    
    If lr = -2 Then
        MsgBox "Ocurrió un error al actualizar el índice del último informe." & vbNewLine & vbNewLine & _
               "Se recomienda no utilizar el programa para prevenir la sobreescritura de datos de otros índices.", vbCritical, "Error al actualizar el índice"
        Exit Sub
    Else
        WriteIni Cfg.Link, "Last", cReport, lr + 1
    End If
End Sub

Function HasReports(ByVal Number As Long) As Boolean
    HasReports = IIf(ReadIni(Cfg.Link, FormatTaskNumber(Number), cReportsQty, 0) > 0, True, False)
End Function

Function ItemUnderMouse(ByVal list_hWnd As Long, ByVal X As Single, ByVal Y As Single) As Long
Dim pt As PointAPI

    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY
    ClientToScreen list_hWnd, pt
    ItemUnderMouse = LBItemFromPt(list_hWnd, pt.X, pt.Y, False)
End Function

Sub DelLastWord(ByRef tb As TextBox)
' Borrar la ultima palabra
Dim WordStart   As String
Dim Trimmed     As String
Dim CurPos      As Long

    With tb
    If Len(.Text) = 0 Then Exit Sub
       
        Replace .Text, Chr(127), ""
       
        CurPos = .SelStart
        Trimmed = Trim(Left(.Text, CurPos))
       
        ' Evitamos un pequeño bug molesto, si Trimmed está vacío suele dar un runtime error 5
        If Trimmed = "" Then Exit Sub
       
        WordStart = InStrRev(Trimmed, Space(1), Len(Trimmed))
       
        .SelStart = WordStart
        .SelLength = CurPos - WordStart
        .SelText = vbNullString
    End With
End Sub

Sub FromWhichForm(ByVal nForm As Byte)
    Select Case nForm
        Case cFromTaskViewer
            fFromFrmTaskViewer = True
            fFromFrmReportViewer = False
        Exit Sub
       
        Case cFromReportViewer
            fFromFrmTaskViewer = False
            fFromFrmReportViewer = True
        Exit Sub
    End Select
End Sub

Function FindItemInComboBox(ByRef cb As ComboBox, ByVal sItem As String) As Integer
Dim iBuffer As Integer: iBuffer = -1
Dim i       As Integer
Dim sBuff() As String


    ReDim sBuff(0 To cb.ListCount - 1)
    
    For i = LBound(sBuff) To UBound(sBuff)
        sBuff(i) = cb.List(i)
       
        If sBuff(i) = sItem Then
            iBuffer = i
            Exit For
        End If
    Next
    
    If iBuffer = -1 Then
        FindItemInComboBox = -1
        Exit Function
    End If
    
    FindItemInComboBox = iBuffer
End Function

Sub ApplyFont(ByVal cCtrl As Control, Optional ByVal FontIndex As Byte = fGlobal)
' --- Configuración forzada ---
' Para mantener la estructura de la interfaz
    If FontIndex = fSidePanel Or FontIndex = fViewers Then
        cCtrl.Font = IIf(fGlobalFont, sFontName(fGlobal), sFontName(FontIndex))
        cCtrl.FontSize = 9
        Exit Sub
    End If
' --- Fin de configuración forzada ---
    
    If fGlobalFont = True Then
        FontIndex = fGlobal
    End If
        
cCtrl.Font = sFontName(FontIndex)
cCtrl.FontSize = sFontSize(FontIndex)
End Sub

Function GetTasksLastReport(ByVal Number As Long) As Long
Dim sIndex As String

    sIndex = ReadIni(Cfg.Link, FormatTaskNumber(Number), cReportsQty, 0)

GetTasksLastReport = CLng(ReadIni(Cfg.Link, FormatTaskNumber(Number), cReport & sIndex, -1))
End Function

Function DisplayTask(ByVal Number As Long, Optional ByVal PadExtraLine As Boolean = False, Optional ByVal Wrap As Boolean = False) As String
Dim pd As Long:     pd = PadMain        ' Pad
Dim nl As String:   nl = vbNewLine      ' NewLine
Dim ol As String                        ' Offset Line
Dim ok As String                        ' Ok Report
Dim ho As String                        ' Has Observations
    
Dim Buffer As String
    
    If Not TaskExists(Number) Then
        DisplayTask = "<El pedido " & FormatTaskNumber(Number) & " no existe.>"
        Exit Function
    End If
    
    Task = ReadTask(Number)
    
    With Task
        If Len(Trim(.Observations)) = 0 Then
            ho = "<Sin observaciones>"
        Else
            ho = .Observations
        End If
           
    ' Dependiendo de los renglones que tenga, añadir espacios para compensar
        Select Case Len(.Observations) / IIf(PadExtraLine, MaxChrWe, MaxChrW)
            Case Is < 1
                ol = nl + nl + nl
            Case Is >= 1
                ol = nl + nl
            Case Is >= 2
            ol = nl
        End Select
    
        If .OkReport <> -1 Then
            ok = " (Informe: " & FormatReportNumber(.OkReport) & ")"
        Else
            ok = ""
        End If
        
        'Debug.Print "Okdate:   " & .OkDate
    
    Buffer = PadStr("Pedido:", pd) & FormatTaskNumber(.Number) & nl & _
             PadStr("Informes:", pd) & .ReportAmount & nl & nl & _
             PadStr("Contacto:", pd) & .Contact & nl & _
             PadStr("Sucursal:", pd) & .BranchOffice & nl & nl & _
             PadStr("Fecha:", pd) & FormatDate2(.Date) & nl & _
             PadStr("Hora:", pd) & FormatTime(.Time) & nl & _
             PadStr("Fecha Ok:", pd) & (IIf(.OkDate = "", strPEND, FormatDate2(.OkDate))) & nl & nl & _
             PadStr("Estado:", pd) & .Status & ok & nl & _
             PadStr("Prioridad:", pd) & .Priority & nl & _
             PadStr("Técnico:", pd) & .Technician & nl & nl & _
             PadStr("--- Observaciones ", IIf(PadExtraLine, MaxChrWe, MaxChrW), "-") & nl & WordWrap(ho, IIf(PadExtraLine, MaxChrWe, MaxChrW)) & ol & _
             PadStr("--- Detalle ", IIf(PadExtraLine, MaxChrWe, MaxChrW), "-") & nl & IIf(Wrap, WordWrap(.Detail, IIf(PadExtraLine, MaxChrWe, MaxChrW)), .Detail) ' .Detail 'WordWrap(.Detail, IIf(PadExtraLine, MaxChrWe, MaxChrW))
        
        DisplayTask = Buffer
    End With
End Function

Function DisplayTask2(ByVal Number As Long) As String
Const nl    As String = vbNewLine

Dim Buffer  As String
Dim pd      As Integer:     pd = 20
Dim lr      As Long:        lr = GetTasksLastReport(Number)

Task = ReadTask(Number)
Report = ReadReport(lr)

    With Task
        Buffer = "*" & String(MaxChrW - 2, "-") & "*" & nl & _
                 "|" & PadStr(" Pedido " & FormatTaskNumber(Number), MaxChrW - 2, " ") & "|" & nl & _
                 "*" & String(MaxChrW - 2, "-") & "*" & nl & _
                 " Contacto: " & .Contact & nl & _
                 PadStr(" Fecha: " & FormatDate2(.Date), pd) & "Fecha Ok: " & FormatDate2(.OkDate) & nl & _
                 PadStr(" Hora:  " & FormatTime(.Time), pd) & "Estado:    " & .Status & nl & nl & _
                 PadStr("--- Detalle ", MaxChrW, "-") & nl & .Detail & nl & nl
    End With

    With Report
        Buffer = Buffer & "*" & String(MaxChrW - 2, "-") & "*" & nl & _
                 "|" & PadStr(" Último informe (" & FormatReportNumber(lr), MaxChrW - 2, " ") & "|" & nl & _
                 "*" & String(MaxChrW - 2, "-") & "*" & nl & _
                 " Contacto: " & .Contact & nl & _
                 " Fecha: " & FormatDate2(.Date) & nl & _
                 " Hora:  " & FormatTime(.Time) & nl & nl & _
                 PadStr("--- Detalle ", MaxChrW, "-") & nl & .Detail & nl & nl
    End With

DisplayTask2 = Buffer
End Function

Function DisplayReport(ByVal Number As Long, Optional ByVal Wrap As Boolean = False) As String
Dim pd  As Long:    pd = PadMain
Dim nl  As String:  nl = vbNewLine

    If Not ReportExists(Number) Then
        DisplayReport = "<El informe " & FormatReportNumber(Number) & " no existe.>"
        Exit Function
    End If
    
   Report = ReadReport(Number)

    With Report
        DisplayReport = PadStr("Informe:", pd) & FormatReportNumber(Number) & nl & _
                        PadStr("Pedido de Ref.:", pd) & FormatTaskNumber(.ReferenceTask) & nl & nl & _
                        PadStr("Fecha:", pd) & FormatDate2(.Date) & nl & _
                        PadStr("Hora:", pd) & FormatTime(.Time) & nl & _
                        PadStr("Contacto:", pd) & .Contact & nl & _
                        PadStr("Sucursal:", pd) & .BranchOffice & nl & nl & _
                        PadStr("Tecnico:", pd) & .Technician & nl & nl & _
                        PadStr("--- Detalle ", MaxChrWe, "-") & nl & IIf(Wrap, WordWrap(.Detail, MaxChrWe), .Detail)  ' .Detail 'WordWrap(.Detail, IIf(PadExtraLine, MaxChrWe, MaxChrW))
    End With
End Function

Function GetMailSubject(ByVal Number As Long, ByVal OfTask As Boolean, Optional ByVal Notify As Boolean = False) As String
Dim sBuffer As String
Dim sDetail As String

    sBuffer = GetDetail(Number, OfTask)
    
    If Len(sBuffer) > SubjectLen Then
        sDetail = Left$(sBuffer, SubjectLen) & "..."
    Else
        sDetail = sBuffer
    End If

    sBuffer = SubjectFormat
    
    sBuffer = Replace(sBuffer, tMailKind, IIf(OfTask, "Pedido", "Informe"))
    sBuffer = Replace(sBuffer, tMailContact, GetContact(Number, OfTask))
    sBuffer = Replace(sBuffer, tMailTechnician, GetTechnician(Number, OfTask))
    sBuffer = Replace(sBuffer, tMailNumber, IIf(OfTask, FormatTaskNumber(Number), FormatReportNumber(Number)))
    sBuffer = Replace(sBuffer, tMailDetail, Replace(sDetail, cNewLine, " "))

    GetMailSubject = sBuffer

    If Notify Then MsgBox "Se copió el contenido en el portapapeles.", vbInformation, "Copiar detalle como asunto"
End Function

Sub NewMail(ByVal frmFrom As Form, ByVal pTo As String, ByVal pSubject As String, ByVal pBody As String)
' Enviar mediante Shell de Windows y mailto:
    ShellExecute frmFrom.hwnd, "open", "mailto:" & pTo & "?subject=" & pSubject & "&body=" & Replace(pBody, vbNewLine, "%0D%0A"), vbNullString, vbNullString, SW_SHOW
End Sub

Public Function WordWrap(ByRef Text As String, ByVal Width As Long, Optional ByRef CountLines As Long) As String
' by Donald, donald@xbeat.net, 20040913
  
Dim i As Long
Dim lenLine As Long
Dim posBreak As Long
Dim cntBreakChars As Long
Dim abText() As Byte
Dim abTextOut() As Byte
Dim ubText As Long

' no fooling around
If Width <= 0 Then
    CountLines = 0
    Exit Function
End If
  
If Len(Text) <= Width Then  ' no need to wrap
    CountLines = 1
    WordWrap = Text
    Exit Function
End If
  
abText = StrConv(Text, vbFromUnicode)
ubText = UBound(abText)
ReDim abTextOut(ubText * 3) 'dim to potential max
  
    For i = 0 To ubText
        Select Case abText(i)
            Case 32, 45 'space, hyphen
            posBreak = i
            
            Case Else
        End Select
        
        abTextOut(i + cntBreakChars) = abText(i)
        lenLine = lenLine + 1
        
        If lenLine > Width Then
            If posBreak > 0 Then
            ' don't break at the very end
                If posBreak = ubText Then Exit For
                ' wrap after space, hyphen
                    abTextOut(posBreak + cntBreakChars + 1) = 13  'CR
                    abTextOut(posBreak + cntBreakChars + 2) = 10  'LF
                    i = posBreak
                    posBreak = 0
                Else
                ' cut word
                    abTextOut(i + cntBreakChars) = 13    'CR
                    abTextOut(i + cntBreakChars + 1) = 10 'LF
                    i = i - 1
                End If
            cntBreakChars = cntBreakChars + 2
            lenLine = 0
        End If
    Next
  
    CountLines = cntBreakChars \ 2 + 1
  
    ReDim Preserve abTextOut(ubText + cntBreakChars)
    WordWrap = StrConv(abTextOut, vbUnicode)
End Function

Function CustomWordWrap(ByVal Text As String, Optional ByVal MaxLineLen As Integer = 70, Optional ByVal AddPipeAtEnd As Boolean = True, Optional ByVal NormalizeHeight As Boolean = True) As String()
Dim i   As Integer          ' Bucle
Dim p() As String           ' Parse
Dim b() As String           ' Buffer

' Si se recibe una cadena vacía
    If Trim(Text) = "" Then
        ReDim b(0)
        b(0) = "| " & PadStr("<El pedido no tiene observaciones.>", MaxLineLen - 4) & " |"
        CustomWordWrap = b
       
        Erase b
       
        Exit Function
    End If
    
    MaxLineLen = MaxLineLen - IIf(AddPipeAtEnd, 4, 2)     ' 4 para "|  |" y 2 para "| "

' Hacer el wrap a la longitud correspondiente y luego guardar en un array partiéndolo por los vbNewLine
    Text = WordWrap(Text, MaxLineLen)
    p = Split(Text, vbNewLine)
    
    ReDim b(LBound(p) To UBound(p))
       
    For i = LBound(b) To UBound(b)
    ' Llenar con informacion mientras el array b sea más pequeño que el array p
        If i <= UBound(p) Then
            b(i) = "| " & PadStr(Trim$(p(i)), MaxLineLen) & IIf(AddPipeAtEnd, " |", "")
        Else
            b(i) = "| " & Space(MaxLineLen) & IIf(AddPipeAtEnd, " |", "")
        End If
    Next i
    
    CustomWordWrap = b
    Erase b
End Function

Function RoundUp(ByVal Number As Single) As Long
' Redondea siempre para arriba
    RoundUp = CLng(Number + 0.5)
End Function

Sub Export(ByVal Tasks As Boolean)
Dim sPrompt As String
Dim sOutput As String
Dim sError  As String

PleaseNotAgain:

sPrompt = InputBox("Ingrese la ruta en la cual quiere guardar el archivo *.csv exportado. Debe tener permisos de escritura sobre esa carpeta.", _
                   "Exportar pedidos", App.Path)

sOutput = "El archivo se guardó correctamente en la ruta %A%." & vbNewLine & vbNewLine & _
          "Los siguientes caracteres fueron reemplazados al generar el archivo:" & vbNewLine & _
          "Comillas (" & Chr(34) & ") por Dobles comillas simples ('')" & vbNewLine & _
          "Punto y coma (;) por Coma (,)" & vbNewLine & _
          "Saltos de línea (CrLf) por Espacio ( )" & vbNewLine & vbNewLine & _
          "¿Desea abrir la ubicación del archivo?"

sError = "Ocurrió un error al guardar el archivo. Verifique que el archivo no está abierto, que tiene acceso a la carpeta y que puede escribir archivos en ella."

    If StrPtr(sPrompt) = 0 Then Exit Sub
    If sPrompt = "" Then Exit Sub

    If Not PathExists(FormatPath(sPrompt)) Then
        MsgBox "La ruta especificada no existe.", vbExclamation
        GoTo PleaseNotAgain
    End If

    If Tasks Then
        If ExportTasks(sPrompt) = True Then
            sOutput = Replace(sOutput, "%A%", strBufferPath)
            If MsgBox(sOutput, vbInformation + vbYesNo) = vbYes Then
                Shell "explorer.exe /select," & FormatPath(strBufferPath), vbNormalFocus
            End If
        Else
            MsgBox sError, vbExclamation
        End If
    Else
        If ExportReports(sPrompt) = True Then
        sOutput = Replace(sOutput, "%A%", strBufferPath)
            If MsgBox(sOutput, vbInformation + vbYesNo) = vbYes Then
                Shell "explorer.exe /select," & FormatPath(strBufferPath), vbNormalFocus
            End If
        Else
            MsgBox sError, vbExclamation
        End If
    End If
End Sub

Sub ExportTask(ByVal Number As Long)
Dim sPrompt As String

PleaseNotAgain:

sPrompt = InputBox("Ingrese la ruta en la cual quiere guardar el archivo *.txt exportado. Debe tener permisos de escritura sobre esa carpeta.", _
                   "Exportar pedidos", _
                   Environ$("TEMP"))

    If StrPtr(sPrompt) = 0 Then Exit Sub
    If sPrompt = "" Then Exit Sub
    
    If Not PathExists(FormatPath(sPrompt)) Then
        MsgBox "La ruta especificada no existe.", vbExclamation
        GoTo PleaseNotAgain
    End If

    If ExportSpecificTask(sPrompt, Number) = True Then
        If MsgBox("El archivo se guardó correctamente en la ruta " & strBufferPath & "." & vbNewLine & vbNewLine & _
                  "¿Desea abrir la ubicación del archivo?", vbInformation + vbYesNo) = vbYes Then
                  Shell "explorer.exe /select," & FormatPath(strBufferPath), vbNormalFocus
        End If
    Else
        MsgBox "Ocurrió un error al guardar el archivo. Verifique que el archivo no está abierto, que tiene acceso a la carpeta y que puede escribir archivos en ella.", vbExclamation
    End If
End Sub

Sub CopyAsSubject(ByVal TaskNumber As Long)
    Clipboard.Clear
    Clipboard.SetText GetMailSubject(TaskNumber, True, fClipboardCopyNotify)
End Sub

Sub CopyToClipboard(ByVal Number As Long, ByVal IsTask As Boolean)
    Clipboard.Clear
    Clipboard.SetText IIf(IsTask, DisplayTask(Number, True, True), DisplayReport(Number, True))
    
    If fClipboardCopyNotify Then
        MsgBox "Se copió el contenido en el portapapeles.", vbInformation, "Copiar pedido en el portapapeles"
    End If
End Sub

Sub SendByMail(ByVal Number As Long, ByVal IsTask As Boolean, ByRef frmFrom As Form)
Dim tSubject    As String
Dim tTo         As String       ' Actualmente en desuso, pero queda por las dudas
Dim tBody       As String

tTo = vbNullString
tSubject = Replace(GetMailSubject(Number, IsTask), cNewLine, " ")

If IsTask Then
    tBody = DisplayTask(Number, True, False)
Else
    tBody = DisplayReport(Number, False)
End If

    NewMail frmFrom, tTo, tSubject, tBody
End Sub

Sub Log(ByVal Message As String)
Dim ff As Integer:  ff = FreeFile
    
    Open App.Path & "\Debug.log" For Append As ff
        Print #ff, "[" & Date & " - " & Time & "] -----------------------------"
        Print #ff, Message
        Print #ff, ""
    Close ff
End Sub

Sub SetFrmMainCaption()
' Se llama desde frmPickTechnician y frmWelcomeSplash
    frmMain.mTechnicianSelect.Caption = "Seleccionar técnico (Actual: " & CurrentTechnician & ")"
    frmMain.Caption = "Sistema de pedidos Lite (Beta) | " & "Técnico actual: " & CurrentTechnician
End Sub

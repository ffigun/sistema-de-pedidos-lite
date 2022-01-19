Attribute VB_Name = "modDeclarations"
Option Explicit

' UDTs
    Public Type ConfigStructure
        Tasks               As String
        Reports             As String
        Link                As String
        Data                As String
        Paths               As String
    End Type

    Public Type TaskStructure
        Number              As Long
     
        Contact             As String       ' * 24
        Detail              As String       ' * 255
        Status              As String
        OkDate              As String * 6   ' ddmmyy
        Date                As String * 6   ' ddmmyy
        Time                As String * 4   ' hhmm
        Observations        As String       ' * 64
        Priority            As String * 1
        BranchOffice        As String
        Technician          As String
        OkReport            As Long
     
        ReportAmount        As Long
    End Type

    Public Type ReportStructure
        Number              As Long
     
        Contact             As String       ' * 64
        Detail              As String       ' * 255
        Date                As String * 6   ' ddmmyy
        Time                As String * 4   ' hhmm
        BranchOffice        As String
        Technician          As String
        ReferenceTask       As Long
    End Type

' Estructuras
    Public Cfg                      As ConfigStructure
    Public Task                     As TaskStructure
    Public Report                   As ReportStructure
    Public TempTask                 As TaskStructure
    Public TempReport               As ReportStructure

' Variables
    Public CurrentTechnician        As String
    Public CurrentTechnicianIndex   As Integer
    Public CurrentTask              As Long
    Public CurrentReport            As Long
    Public TaskHolder               As Long                     ' Antes de entrar a FrmSearch, guardar el valor de CurrentReport provisoriamente
    Public MainFilter               As Byte
    Public strBufferPath            As String
    Public ItemHover                As Long
    Public frmMainW                 As Long
    Public frmMainH                 As Long
    Public TechnicianAmount         As Integer
    Public sFontName(4)             As String
    Public sFontSize(4)             As Integer
    Public WHThreshold              As Single                   ' % de tamaño de pantalla que no debe superar ni Width ni Height
    Public gSepChar                 As Integer                  ' GlobalSepChar, cantidad de espacio entre campos
    Public SubjectLen               As Integer                  ' Cantidad de caracteres a tomar del pedido
    Public SubjectFormat            As String                   ' El formato que tiene que tener el asunto del mail

    Public WikiPath                 As String
    Public BackupsPath              As String
    Public MaxBackups               As Byte                     ' Cantidad de archivos de backup que se crearan

    Public Msg                      As String
    Public Tip                      As Integer
    Public fFilterMore              As Boolean

    Public CurrentFilter            As Integer
    Public sType                    As String
    Public sTypes                   As String
    Public sParse()                 As String
    Public sSearch(1 To 15)         As String
    Public ExportedFileName         As String
    Public TaskFileName             As String
    Public lstCurrent               As ListBox

' Constantes
    Public Const MaxCharDetail      As Integer = 255            ' Tecnicamente soportan caracteres ilimitados. Pero se fuerza a 255
    Public Const wOffset            As Integer = 240            ' Cantidad de twips que representan 1 unidad de cWidth
    Public Const hOffset            As Integer = 240            ' Cantidad de twius que representan 1 unidad de cHeight
    Public Const MaxChrW            As Integer = 40             ' Cantidad de caracteres de cWidth del txtDetalles (Fixed-Width font)
    Public Const MaxChrWe           As Integer = 48             ' Idem MaxChrW pero extendido, para frmtaskviewer y frmReportViewer
    Public Const PadMain            As Integer = 17             ' Pad que separa las columnas de los valores

    Public Const strOK              As String = "OK"            ' String para OK
    Public Const strPEND            As String = "PEND"          ' String para PEND

' Campos
    Public Const UBoundTaskData = 10
    Public Enum TaskDataIndex
        Contact = 0
        Detail = 1
        Status = 2
        OkDate = 3
        InputDate = 4
        InputTime = 5
        Observations = 6
        Priority = 7
        BranchOffice = 8
        Technician = 9
        OkReport = 10
    End Enum
    
    Public Const UBoundReportData = 6
    Public Enum ReportDataIndex
        Contact = 0
        Detail = 1
        InputDate = 2
        InputTime = 3
        BranchOffice = 4
        Technician = 5
        ReferenceTask = 6
    End Enum
   
' Campos de archivos INI
    Public Const cData              As String * 4 = "DATA"
    Public Const cReportsQty        As String * 2 = "QT"
    Public Const cReport            As String * 2 = "RP"
    Public Const cTask              As String * 2 = "TK"
    
' Exportación de archivos
    Public Const cTasks             As Boolean = True
    Public Const cReports           As Boolean = False
    
' Modos de reparar interfaz busqueda
    Public Const cWidth             As Byte = 0
    Public Const cHeight            As Byte = 1
    Public Const cBoth              As Byte = 2
    
' Filtrar por estado de pedido
    Public Const cPEND              As Byte = 1
    Public Const cOK                As Byte = 2
    Public Const cTODOS             As Byte = 3

' Modos de mostrar los informes
    Public Const mAll               As Byte = 1
    Public Const mSpecific          As Byte = 2

' Eventos personalizados segun desde que ventana se llaman
    Public Const cFromTaskViewer    As Byte = 1
    Public Const cFromReportViewer  As Byte = 2

' Indices del array de fuentes
    Public Const fMain              As Byte = 0
    Public Const fSidePanel         As Byte = 1
    Public Const fViewers           As Byte = 2
    Public Const fSearch            As Byte = 3
    Public Const fGlobal            As Byte = 4

' Búsqueda
    Public Const csNumber           As Byte = 1
    Public Const csDetail           As Byte = 2
    Public Const csTechnician       As Byte = 3
    Public Const csBranchOffice     As Byte = 4
    Public Const csContact          As Byte = 5
    Public Const csDate             As Byte = 6
    Public Const csOkDate           As Byte = 7
    Public Const csStatus           As Byte = 8
    Public Const csPriority         As Byte = 9
    Public Const csObservations     As Byte = 10
    Public Const csHasObservations  As Byte = 11

    Public Const csByTask           As Byte = 0
    Public Const csByReport         As Byte = 1

    Public Const csMsgInformation   As Byte = 0
    Public Const csMsgWarning       As Byte = 1
    Public Const csMsgResult        As Byte = 2
    Public Const csMsgNone          As Byte = 9

' Códigos de escape y comodines
    Public Const cNewLine           As String = "{\n}"          ' Se interpreta como un salto de linea
    Public Const cSepChar           As String = "{[^]}"         ' Separa los campos al leer el Key DATA= de los INIs (INIs V2)
    Public Const cAppPath           As String = "&f"            ' Se interpreta como App.Path
    Public Const tMailKind          As String = "%K"            ' Se interpreta como el tipo (Pedido / Informe)
    Public Const tMailNumber        As String = "%N"            ' Se interpreta como el número
    Public Const tMailDetail        As String = "%D"            ' Se interpreta como el detalle
    Public Const tMailContact       As String = "%C"            ' Se interpreta como el contacto
    Public Const tMailTechnician    As String = "%T"            ' Se interpreta como el tecnico

' Antonio Flags
    Public fNewReport               As Boolean      ' Si es un nuevo informe, personalizar el form, de lo contrario, cargar los datos para editar
    Public fNewTask                 As Boolean      ' Si es un nuevo pedido, personalizar el form, de lo contrario, cargar los datos para editar
    Public fShowSpecificReport      As Boolean      ' Indica que se exige a frmReportViewer que muestre un informe especifico
    Public fSearchByTask            As Boolean      ' Si es true, se busca por pedidos, de lo contrario se busca por informe
    Public fShowPriority            As Boolean      ' Se alza cuando se usa la opcion de mostrar la prioridad de los pedidos
    Public fHlTasksWithReports      As Boolean      ' Se alza cuando se usa la opcion de resaltar pedidos con informes
    Public fShowWiki                As Boolean      ' Permite ver la opcion para usar la Wiki
    Public fUseBuiltInBackup        As Boolean      ' Permite usar backups integrados
    Public fClipboardCopyNotify     As Boolean      ' Muestra un MsgBox avisando que se realizo la copia de la informacion en el portapapeles
    Public fAlreadySetFocus         As Boolean      ' Workaround para que se realice el focus a un textbox mientras se dibujan los forms
    Public fSearchFilterOn          As Boolean      ' Indica que se debe buscar sobre los resultados de busqueda
    Public fGlobalFont              As Boolean      ' Especifica si utilizar una única fuente para mostrar la información formateada
    Public fAllowSearch             As Boolean      ' Impide que se busque mientras se esta buscando
    Public fFromFrmNewTask          As Boolean      ' Indica que se llamo el evento desde frmNewTask
    Public fFromFrmTaskViewer       As Boolean      ' Indica que se llamo el evento desde frmTaskViewer
    Public fFromFrmReportViewer     As Boolean      ' Indica que se llamo el evento desde frmReportViewer
    Public fFromFrmMain             As Boolean      ' Indica que se llamo el evento desde frmMain
    Public fFromFrmSearch           As Boolean      ' Indica que se llamo el evento desde frmSearch
    Public fAbortBackup             As Boolean      ' Aborta el proceso de backup si se encuentra una anomalía
    Public fFoundSomething          As Boolean      ' Indica que se encontró al menos un item durante la búsqueda
    Public fParameter2              As Boolean      ' Indica si hay un segundo parámetro activo en frmSearch
    Public fNotNumeric              As Boolean      ' Indica que alguno de los valores comparados no es numerico
    Public fDebug                   As Boolean      ' Para obtener datos fuera del IDE
    Public fDisplayTaskAndReport    As Boolean      ' Muestra el pedido y el ultimo informe del mismo en frmMain
'   Public fTaskWasOk               As Boolean      ' Indica si el pedido estaba OK cuando se comenzó a editarlo

' *-----------------------------*
' |     M O D O   D E B U G     |
' *-----------------------------*
' Public Const fDebug               As Boolean = False

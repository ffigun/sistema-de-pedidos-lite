VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Buscar"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraResults 
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   6495
      Begin VB.Timer tmrSearchDelay 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4440
         Top             =   3960
      End
      Begin VB.CheckBox chkFilterMore 
         Caption         =   "Buscar sobre este resultado"
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
         TabIndex        =   13
         Top             =   4320
         Width           =   6255
      End
      Begin VB.ListBox lstTasks 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         IntegralHeight  =   0   'False
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.ListBox lstReports 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   6255
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2535
      Begin VB.CommandButton cmdGo 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cmbParameter1 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtParameter2 
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtParameter1 
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
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblParameter2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblParameter1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox picAntibug 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2295
         TabIndex        =   14
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton optSearchBy 
            Caption         =   "Pedidos"
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
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton optSearchBy 
            Caption         =   "Informes"
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
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.ComboBox cmbSearch 
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
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblDetail 
      Caption         =   "Holis"
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
      TabIndex        =   12
      Top             =   4920
      Width           =   9135
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FormH   As Long         ' Form Height
Dim FormW   As Long         ' Form Width

Dim FormSH  As Long         ' Form ScaleHeight
Dim FormSW  As Long         ' Form ScaleWidth

Private Sub chkFilterMore_Click()
' Filtrar sobre el resultado de busqueda
    If chkFilterMore.Value = vbChecked Then
        fFilterMore = True
        optSearchBy(csByTask).Enabled = False
        optSearchBy(csByReport).Enabled = False
    Else
        fFilterMore = False
        optSearchBy(csByTask).Enabled = True
        optSearchBy(csByReport).Enabled = True
    End If
End Sub

Private Sub cmbParameter1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdGo_Click
    End If
End Sub

Private Sub cmbSearch_Click()
' El evento _Click es cuando cambia, sin importar como.

' Si cambia de filtro de busqueda, vaciar las cajas de texto
    If Not CurrentFilter = cmbSearch.ListIndex + 1 Then       ' Uso el +1 porque .ListIndex es base 0
        txtParameter1.Text = ""
        txtParameter2.Text = ""
             
    ' Solo limpiar las listas si no se esta filtrando su contenido
        If Not fFilterMore Then
            lstReports.Clear
            lstTasks.Clear
            chkFilterMore.Enabled = False
            chkFilterMore.Value = vbUnchecked
        End If
    End If

' Formatear interfaz
    With txtParameter1
    
        If cmbSearch.ListIndex + 1 <> csBranchOffice Or cmbSearch.ListIndex + 1 <> csTechnician Then
            .Visible = True
            cmbParameter1.Visible = False
        End If
        
        lblParameter1.Caption = ""
        lblParameter2.Caption = ""
        
        Select Case (cmbSearch.ListIndex + 1)
            Case csNumber
                DisableInput 1, False
                DisableInput 2, False
                    
                .MaxLength = 5
                txtParameter2.MaxLength = 5
                lblParameter1.Caption = "Desde:"
                lblParameter2.Caption = "Hasta:"
                
                CurrentFilter = csNumber
            GoTo ExitCase
              
            Case csDetail
                DisableInput 1, False
                DisableInput 2, True
                
                .MaxLength = 0
                lblParameter1.Caption = "Palabra clave:"
                
                CurrentFilter = csDetail
            GoTo ExitCase
            
            Case csTechnician
                DisableInput 2, True
                
                FillParametersComboBox cmbParameter1, csTechnician
                .Visible = False
                cmbParameter1.Visible = True
                lblParameter1.Caption = "Técnico:"
                
                CurrentFilter = csTechnician
            GoTo ExitCase
            
            Case csBranchOffice
                DisableInput 2, True
              
                FillParametersComboBox cmbParameter1, csBranchOffice
                .Visible = False
                cmbParameter1.Visible = True
                lblParameter1.Caption = "Sucursal:"
                
                CurrentFilter = csBranchOffice
            GoTo ExitCase
            
            Case csContact
                DisableInput 1, False
                DisableInput 2, True
                
                .MaxLength = 0
                lblParameter1.Caption = "Contacto:"
                
                CurrentFilter = csContact
            GoTo ExitCase
            
            Case csDate
                DisableInput 1, False
                DisableInput 2, False
                
                .MaxLength = 6
                txtParameter2.MaxLength = 6
                lblParameter1.Caption = "Desde:"
                lblParameter2.Caption = "Hasta:"
                
                CurrentFilter = csDate
            GoTo ExitCase
            
            Case csOkDate
                DisableInput 1, False
                DisableInput 2, False
                
                .MaxLength = 6
                txtParameter2.MaxLength = 6
                lblParameter1.Caption = "Desde:"
                lblParameter2.Caption = "Hasta:"
                
                CurrentFilter = csOkDate
            GoTo ExitCase
            
            Case csStatus
                DisableInput 2, True
                
                FillParametersComboBox cmbParameter1, csStatus
                .Visible = False
                cmbParameter1.Visible = True
                lblParameter1.Caption = "Estado:"
                
                CurrentFilter = csStatus
            GoTo ExitCase
            
            Case csPriority
                DisableInput 2, True
                
                FillParametersComboBox cmbParameter1, csPriority
                .Visible = False
                cmbParameter1.Visible = True
                lblParameter1.Caption = "Prioridad:"
                
                CurrentFilter = csPriority
            GoTo ExitCase
            
            Case csObservations
                  DisableInput 1, False
                  DisableInput 2, True
                
                  .MaxLength = 0
                  lblParameter1.Caption = "Palabra clave:"
                  CurrentFilter = csObservations
            GoTo ExitCase
            
            Case csHasObservations
                DisableInput 1, True
                DisableInput 2, True
            
                CurrentFilter = csHasObservations
            GoTo ExitCase
        End Select
    
ExitCase:
        ShowMessages
    End With
End Sub

Private Sub cmdGo_Click()
' Antispam
    If fAllowSearch = False Then Exit Sub

Dim sErr        As String
Dim sFiltered   As String       ' Almacena los pedidos filtrados, separados por un pipe (|)
Dim i           As Long         ' Bucle
Dim j           As Long         ' Bucle

Dim lStart      As Long
Dim lEnd        As Long
Dim LastTorR    As Long         ' Último Pedido o Informe

' Búfferes de distintos tipos
Dim sBuffer     As String
Dim sBuffer2    As String
Dim iBuffer     As Long
Dim lBuffer     As Long

Dim dDate       As Date
Dim dFrom       As Date
Dim dTo         As Date

' Deshabilitar búsqueda para evitar el spam
    Screen.MousePointer = vbHourglass
    cmdGo.Enabled = False

' Bajar flags correspondientes, definidas en modDeclarations
    fFoundSomething = False
    fAllowSearch = False
    fParameter2 = False
    
    If Not fFilterMore Then
        lstTasks.Clear
        lstReports.Clear
    End If
        
    If fSearchByTask Then
        Set lstCurrent = lstTasks
    Else
        Set lstCurrent = lstReports
    End If

' Validar datos genericos
    If fSearchByTask Then
        If LastTask = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If
    Else
        If LastReport = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If
    End If
    
    If txtParameter2.Enabled = True Then
        If Trim(txtParameter2.Text) <> "" Then
            fParameter2 = True
        End If
    End If
    
' Ultimo Pedido o Informe, sirve para no tener que hacer una busqueda para informes y otra para pedidos
    LastTorR = IIf(fSearchByTask, LastTask, LastReport)

' Buscar
    Select Case CurrentFilter
        Case csNumber ' ------------------------------------------------------------------------------------------
        ' Validar
        If Trim(txtParameter2.Text) <> "" Then
            If Not IsNumeric(txtParameter1.Text) Or Not IsNumeric(txtParameter2.Text) Then
                sErr = "Debe ingresar un número de " & sType & " válido."
                GoTo OhNo
            End If
          
            If Val(txtParameter1.Text) < 0 Or Val(txtParameter2.Text) < 0 Then
                sErr = "El número de " & sType & " no puede ser negativo."
                GoTo OhNo
            End If
            
            If Val(txtParameter1.Text) > Val(txtParameter2.Text) Then
                sErr = "El pedido inicial no puede ser mayor que el lEndal."
                GoTo OhNo
            End If
        Else
            If Not IsNumeric(txtParameter1.Text) Then
                sErr = "Debe ingresar un número de " & sType & " válido."
                GoTo OhNo
            End If
            
            If Val(txtParameter1.Text) < 0 Then
                sErr = "El número de " & sType & " no puede ser negativo."
                GoTo OhNo
            End If
        End If
          
        ' Formatear
        If fSearchByTask Then
            txtParameter1.Text = FormatTaskNumber(txtParameter1.Text)
            If txtParameter2.Text <> "" Then txtParameter2.Text = FormatTaskNumber(txtParameter2.Text)
        Else
            txtParameter1.Text = FormatReportNumber(txtParameter1.Text)
            If txtParameter2.Text <> "" Then txtParameter2.Text = FormatReportNumber(txtParameter2.Text)
        End If
          
        ' Buscar por pedido
        If Not fParameter2 Then
        ' Sólo 1 campo escrito
            sFiltered = "|" & Val(txtParameter1.Text)
            
            If fFilterMore Then
            ' Buscar sobre resultados
                For i = 0 To lstCurrent.ListCount
                    If fSearchByTask Then
                        If Mid$(lstCurrent.List(i), 1, 5) = txtParameter1.Text Then
                            AddTask (sFiltered)
                            fFoundSomething = True
                            Exit For
                        End If
                    Else
                        If Mid$(lstCurrent.List(i), 1, 5) = txtParameter1.Text Then
                            AddReport (sFiltered)
                            fFoundSomething = True
                            Exit For
                        End If
                    End If
                Next
            Else
        ' Busqueda comun
            If fSearchByTask Then
                If TaskExists(Val(txtParameter1.Text)) Then
                    AddTask (sFiltered)
                        fFoundSomething = True
                    Else
                        fFoundSomething = False
                    End If
                Else
                    If ReportExists(Val(txtParameter1.Text)) Then
                        AddReport (sFiltered)
                        fFoundSomething = True
                    Else
                        fFoundSomething = False
                    End If
                End If
            End If
        Else
        ' 2 campos escritos
            If fFilterMore Then
            ' Buscar sobre resultados
                
                lStart = Val(txtParameter1.Text)
                lEnd = Val(txtParameter2.Text)
                    
                For i = 0 To lstCurrent.ListCount - 1
                iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                    If iBuffer >= lStart And iBuffer <= lEnd Then
                        sFiltered = sFiltered & "|" & iBuffer
                        fFoundSomething = True
                    End If
                Next
            Else
            ' Busqueda comun
            ' Buscar como máximo hasta el último pedido / informe
                If Val(txtParameter2.Text) > LastTorR Then
                    j = LastTorR
                Else
                    j = Val(txtParameter2.Text)
                End If
            
                For i = Val(txtParameter1.Text) To j
                    If fSearchByTask Then
                    ' Por pedido
                        If TaskExists(i) Then
                            sFiltered = sFiltered & "|" & i
                            fFoundSomething = True
                        End If
                    Else
                    ' Por informe
                        If ReportExists(i) Then
                            sFiltered = sFiltered & "|" & i
                            fFoundSomething = True
                        End If
                    End If
                Next
            End If
        End If
        
        GoTo ExitCase
        
    Case csDetail ' ------------------------------------------------------------------------------------------
    ' Validar
        If Trim(txtParameter1.Text) = "" Then
            fFoundSomething = False
            GoTo ExitCase
        End If
        
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If InStr(1, LCase(GetDetail(iBuffer, fSearchByTask)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If InStr(1, LCase(GetDetail(i, fSearchByTask)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & i
                         fFoundSomething = True
              End If
            Next
        End If
              
      GoTo ExitCase
                    
    Case csTechnician ' ------------------------------------------------------------------------------------------
    ' Validar
        If cmbParameter1.ListIndex = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If
        
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If UCase(GetTechnician(iBuffer, fSearchByTask)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If UCase(GetTechnician(i, fSearchByTask)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
        
        GoTo ExitCase
        
    Case csBranchOffice ' ------------------------------------------------------------------------------------------
    ' Validar
        If cmbParameter1.ListIndex = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If
        
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If UCase(GetBranchOffice(iBuffer, fSearchByTask)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If UCase(GetBranchOffice(i, fSearchByTask)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
        
        GoTo ExitCase
        
    Case csContact ' ------------------------------------------------------------------------------------------
    ' Validar
        If Trim(txtParameter1.Text) = "" Then
            fFoundSomething = False
            GoTo ExitCase
        End If
        
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If InStr(1, LCase(GetContact(iBuffer, fSearchByTask)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If InStr(1, LCase(GetContact(i, fSearchByTask)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
        
        GoTo ExitCase
        
    Case csDate ' ------------------------------------------------------------------------------------------
        ' Si no escribe nada, tomar la fecha del dia
        If txtParameter1.Text = "" Then
            txtParameter1.Text = Format(Date, "ddmmyy")
        End If
        
        ' Validar
        If Not IsDateValid(txtParameter1.Text) Then
            sErr = "La dDate ingresada en el primer campo no es válida."
            GoTo OhNo
        End If
        
        If fParameter2 Then
            If Not IsDateValid(txtParameter2) Then
                sErr = "La dDate ingresada en el segundo campo no es válida."
                GoTo OhNo
            End If
        End If
        
        dFrom = FormatDate(txtParameter1.Text)
        If fParameter2 Then dTo = FormatDate(txtParameter2.Text)
        
        If fFilterMore Then
        ' Buscar sobre resultados
            If fParameter2 Then
            ' Entre dos fechas
                For i = 0 To lstCurrent.ListCount - 1
                iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                dDate = GetDate2(iBuffer, fSearchByTask)
                
                    If dDate >= dFrom And dDate <= dTo Then
                        sFiltered = sFiltered & "|" & iBuffer
                        fFoundSomething = True
                    End If
                Next
            Else
            ' Solo los de una fecha
                For i = 0 To lstCurrent.ListCount - 1
                iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                dDate = GetDate2(iBuffer, fSearchByTask)
                
                    If dDate = dFrom Then
                        sFiltered = sFiltered & "|" & iBuffer
                        fFoundSomething = True
                    End If
                Next
            End If
        Else
            ' Busqueda comun
            For i = 0 To LastTorR
                dDate = GetDate2(i, fSearchByTask, False)
                
                If fParameter2 Then
                ' Entre dos fechas
                    If dDate >= dFrom And dDate <= dTo Then
                        sFiltered = sFiltered & "|" & i
                        fFoundSomething = True
                    End If
                      Else
              ' Solo los de una fecha
                    If dDate = dFrom Then
                        sFiltered = sFiltered & "|" & i
                        fFoundSomething = True
                    End If
                End If
            Next
        End If
        
        GoTo ExitCase
        
    Case csOkDate ' ------------------------------------------------------------------------------------------
        ' Si no escribe nada, tomar la fecha del dia
        If txtParameter1.Text = "" Then
            txtParameter1.Text = Format(Date, "ddmmyy")
        End If
        
        ' Validar
        If Not IsDateValid(txtParameter1.Text) Then
            sErr = "La dDate ingresada en el primer campo no es válida."
            GoTo OhNo
        End If
        
        If fParameter2 Then
            If Not IsDateValid(txtParameter2) Then
                sErr = "La dDate ingresada en el segundo campo no es válida."
                GoTo OhNo
            End If
        End If
        
        dFrom = FormatDate(txtParameter1.Text)
        If fParameter2 Then dTo = FormatDate(txtParameter2.Text)
        
        If fFilterMore Then
        ' Buscar sobre resultados
            If fParameter2 Then
            ' Entre dos fechas
                For i = 0 To lstCurrent.ListCount - 1
                iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                dDate = GetDate2(iBuffer, fSearchByTask, True)
                
                    If dDate >= dFrom And dDate <= dTo Then
                        sFiltered = sFiltered & "|" & iBuffer
                        fFoundSomething = True
                    End If
                Next
            Else
            ' Solo los de una fecha
                For i = 0 To lstCurrent.ListCount - 1
                iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                dDate = GetDate2(iBuffer, fSearchByTask, True)
                
                    If dDate = dFrom Then
                        sFiltered = sFiltered & "|" & iBuffer
                        fFoundSomething = True
                    End If
                Next
            End If
        Else
            ' Busqueda comun
            For i = 0 To LastTorR
                If GetStatus(i) = strOK Then
                    dDate = GetDate2(i, fSearchByTask, True)
                    
                    If fParameter2 Then
                    ' Entre dos fechas
                        If dDate >= dFrom And dDate <= dTo Then
                            sFiltered = sFiltered & "|" & i
                            fFoundSomething = True
                        End If
                    Else
                    ' Solo los de una dDate
                        If dDate = dFrom Then
                            sFiltered = sFiltered & "|" & i
                            fFoundSomething = True
                        End If
                    End If
                End If
            Next
        End If
          
        GoTo ExitCase

    Case csStatus ' ------------------------------------------------------------------------------------------
    ' Validar
        If cmbParameter1.ListIndex = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If
        
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If UCase(GetStatus(iBuffer)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If UCase(GetStatus(i)) = UCase(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
        
        GoTo ExitCase

    Case csPriority ' ------------------------------------------------------------------------------------------
    ' Validar
        If cmbParameter1.ListIndex = -1 Then
            fFoundSomething = False
            GoTo ExitCase
        End If

        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If GetPriority(iBuffer) = Val(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If GetPriority(i) = Val(cmbParameter1.Text) Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
              
    GoTo ExitCase
                  
    Case csObservations ' ------------------------------------------------------------------------------------------
    ' No hace falta validar nada
    
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If InStr(1, LCase(GetObservations(iBuffer)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If InStr(1, LCase(GetObservations(i)), LCase(txtParameter1.Text)) > 0 Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
        
        GoTo ExitCase
        
    Case csHasObservations ' ------------------------------------------------------------------------------------------
    ' No hace falta validar nada
    
        If fFilterMore Then
        ' Buscar sobre resultados
            For i = 0 To lstCurrent.ListCount - 1
            iBuffer = Val(Mid$(lstCurrent.List(i), 1, 5))
                If Trim(GetObservations(iBuffer)) <> "" Then
                    sFiltered = sFiltered & "|" & iBuffer
                    fFoundSomething = True
                End If
            Next
        Else
        ' Busqueda comun
            For i = 0 To LastTorR
                If Trim(GetObservations(i)) <> "" Then
                    sFiltered = sFiltered & "|" & i
                    fFoundSomething = True
                End If
            Next
        End If
          
        GoTo ExitCase
          
End Select

ExitCase:
    If fFoundSomething = False Then
        Message "No se encontraron resultados con los criterios seleccionados.", csMsgWarning
                  
        If (Not fFilterMore) Or (lstCurrent.ListCount = 0) Then
            chkFilterMore.Value = vbUnchecked
            chkFilterMore.Enabled = False
            fFilterMore = False
        End If
        
        ' Cursor normal
        Screen.MousePointer = vbDefault
        cmdGo.Enabled = True
        fAllowSearch = True
        Exit Sub
    End If
    
' Mostrar los resultados de busqueda
    If fSearchByTask Then
        AddTask sFiltered
    Else
        AddReport sFiltered
    End If

    sParse() = Split(sFiltered, "|")
    If UBound(sParse) = 1 Then
        Message "Se encontró " & 1 & " resultado.", csMsgResult
    Else
        Message "Se encontraron " & UBound(sParse) & " resultados.", csMsgResult
    End If
    
    chkFilterMore.Enabled = True
    
' Cursor normal
    Screen.MousePointer = vbDefault
    cmdGo.Enabled = True
    tmrSearchDelay.Enabled = True
Exit Sub
    
OhNo:
    MsgBox sErr, vbExclamation, "Resuelva el siguiente problema"
    chkFilterMore.Enabled = False

    ' Cursor normal
        Screen.MousePointer = vbDefault
        cmdGo.Enabled = True
        tmrSearchDelay.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Guardar antes de salir
    WriteIni Cfg.Data, "SEARCH", "frmW", Me.Width
    WriteIni Cfg.Data, "SEARCH", "frmH", Me.Height
    
' Limpiar los actuales antes de salir
'  TaskHolder = TaskHolder
    CurrentReport = -1
    frmMain.Update
     
    fFilterMore = False
    fFromFrmSearch = False
End Sub

Private Sub Form_Resize()
' Impedir que el tamaño baje de 9600 x 5820 Twips
    If Me.Width < 9600 Or Me.Width > Screen.Width * WHThreshold Then
        ResetUI cWidth
        Exit Sub
    End If
    
    If Me.Height < 5820 Or Me.Height > Screen.Height * WHThreshold Then
        ResetUI cHeight
        Exit Sub
    End If

    FormSH = Me.ScaleHeight
    FormSW = Me.ScaleWidth

' Relativos a Form
    fraResults.Move 2760, 120, FormSW - fraResults.Left - 120, FormSH - 600
    lblDetail.Move 120, FormSH - lblDetail.Height, FormSW - 240

' Relativos a Frames
    lstReports.Move 120, 300, fraResults.Width - 240, fraResults.Height - 735
    lstTasks.Move 120, 300, fraResults.Width - 240, fraResults.Height - 735
    chkFilterMore.Move 120, fraResults.Height - 375
End Sub

Private Sub Form_Load()
' Interfaz
    FormH = CLng(ReadIni(Cfg.Data, "SEARCH", "frmH", 5820))
    FormW = CLng(ReadIni(Cfg.Data, "SEARCH", "frmW", 9600))
    
    Me.Height = FormH
    Me.Width = FormW
    
    ApplyFont lstTasks, fSearch
    ApplyFont lstReports, fSearch
    
' Limpiar
    lblParameter1.Caption = "Dato 1:"
    txtParameter1.Text = ""
    DisableInput 1, True
    
    lblParameter2.Caption = "Dato 2:"
    txtParameter2.Text = ""
    DisableInput 2, True
    
' Asignar valores
    sSearch(1) = "Número"
    sSearch(2) = "Detalle"
    sSearch(3) = "Técnico"
    sSearch(4) = "Sucursal"
    sSearch(5) = "Contacto"
    sSearch(6) = "Fecha"
    sSearch(7) = "Fecha Ok"
    sSearch(8) = "Estado"
    sSearch(9) = "Prioridad"
    sSearch(10) = "Observaciones"
    sSearch(11) = "Tiene Obs."

    lstTasks.Move lstReports.Left, lstReports.Top, lstReports.Width, lstReports.Height
    lstTasks.Visible = False
    lstReports.Visible = False
    cmbParameter1.Move txtParameter1.Left, txtParameter1.Top, txtParameter1.Width
    fFilterMore = False
    chkFilterMore.Enabled = False
    fAllowSearch = True
    
    sType = IIf(fSearchByTask, "pedido", "informe")
    sTypes = IIf(fSearchByTask, "pedidos", "informes")
     
    lstReports.Clear
    lstReports.Refresh
     
    optSearchBy(csByTask).Value = vbUnchecked
    Call optSearchBy_Click(csByTask)
    optSearchBy(csByTask).Value = vbChecked
    fSearchByTask = True
    fFromFrmSearch = True
    
    If fDebug Then
    Log "frmSearch was called. Flags are as follow:" & vbNewLine & _
        "fFilterMore is " & fFilterMore & vbNewLine & _
        "fFromFrmMain is " & fFromFrmMain & vbNewLine & _
        "fFromFrmReportViewer is " & fFromFrmReportViewer & vbNewLine & _
        "fFromFrmTaskViewer is " & fFromFrmTaskViewer & vbNewLine & _
        "fFromFrmSearch is " & fFromFrmSearch & vbNewLine & _
        "fSearchByTask is " & fSearchByTask
    End If
End Sub

Private Sub lstReports_DblClick()
    fFromFrmSearch = True
    fShowSpecificReport = True
    CurrentReport = Val(Mid$(lstReports.List(lstReports.ListIndex), 1, 5))
     
    frmReportViewer.Show vbModal, Me
End Sub

Private Sub lstReports_KeyPress(KeyAscii As Integer)
    If lstReports.ListCount = 0 Then Exit Sub
    If lstReports.ListIndex = -1 Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call lstReports_DblClick
    End If
End Sub

Private Sub lstTasks_DblClick()
    TaskHolder = Val(Mid$(lstTasks.List(lstTasks.ListIndex), 1, 5))
    frmTaskViewer.Show vbModal, Me
End Sub

Private Sub lstTasks_KeyPress(KeyAscii As Integer)
    If lstTasks.ListCount = 0 Then Exit Sub
    If lstTasks.ListIndex = -1 Then Exit Sub

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call lstTasks_DblClick
    End If
End Sub

Private Sub optSearchBy_Click(Index As Integer)
Dim iTemp As Integer: iTemp = cmbSearch.ListIndex

' Impedir que carguen dos veces lo mismo
If optSearchBy(Index).Value = vbChecked Then Exit Sub

    FillFilterComboBox Index, iTemp

    If Index = csByTask Then
        lstReports.Clear
        
        fSearchByTask = True
        lstTasks.Visible = True
        lstReports.Visible = False
    Else
        lstTasks.Clear
         
        fSearchByTask = False
        lstTasks.Visible = False
        lstReports.Visible = True
    End If
     
    sType = IIf(fSearchByTask, "pedido", "informe")
    sTypes = IIf(fSearchByTask, "pedidos", "informes")
    
    chkFilterMore.Enabled = False
    chkFilterMore.Value = vbUnchecked
     
    ShowMessages
End Sub

Private Sub tmrSearchDelay_Timer()
' Se llama una vez que se bloquea el botón de búsqueda. Luego de un "tick" rehabilita la función de buscar
    fAllowSearch = True
    tmrSearchDelay.Enabled = False
End Sub

Private Sub txtParameter1_GotFocus()
    txtParameter1.SelStart = 0
    txtParameter1.SelLength = Len(txtParameter1.Text)
End Sub

Private Sub txtParameter1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    ' Si el textbox esta vacio y esta buscando por fecha, poner la fecha de hoy
        If (cmbSearch.ListIndex + 1) = csDate Or (cmbSearch.ListIndex + 1) = csOkDate Then
            If Trim(txtParameter1.Text) = "" Then
            ' Formatear dd/mm/yyyy a ddmmyy
                txtParameter1.Text = Format(CStr(Day(Date)) & "/" & CStr(Month(Date)) & "/" & CStr(Year(Date)), "ddmmyy")
            End If
        End If
        
        If fAllowSearch = False Then Exit Sub
        Call cmdGo_Click
    End If
End Sub

Private Sub txtParameter2_GotFocus()
    txtParameter2.SelStart = 0
    txtParameter2.SelLength = Len(txtParameter2.Text)
End Sub

Private Sub txtParameter1_LostFocus()
' Formatear el texto
    With txtParameter1
        Select Case CurrentFilter
            Case csNumber
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
              
                If Not IsNumeric(.Text) Then
                    Message "El número especificado en el campo 1 no es válido.", csMsgWarning
                    Exit Sub
                End If
              
                If fSearchByTask Then
                    .Text = FormatTaskNumber(Val(.Text))
                Else
                    .Text = FormatReportNumber(Val(.Text))
                End If
                  
                Message Msg, Tip
            Exit Sub
                  
            Case csDate
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
              
                If Not IsNumeric(.Text) Then
                    Message "La fecha especificada en el campo 1 no es válida.", csMsgWarning
                    Exit Sub
                End If
              
                If Not IsDateValid(.Text) Then
                    Message "La fecha especificada en el campo 1 no corresponde a un día existente.", csMsgWarning
                    Exit Sub
                End If
                  
                Message Msg, Tip
            Exit Sub
                
            Case csOkDate
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
              
                If Not IsNumeric(.Text) Then
                    Message "La fecha especificada en el campo 1 no es válida.", csMsgWarning
                    Exit Sub
                End If
              
                If Not IsDateValid(.Text) Then
                    Message "La fecha especificada en el campo 1 no corresponde a un día existente.", csMsgWarning
                    Exit Sub
                End If
              
                If IsDateValid(txtParameter2.Text) Then
                    If FormatDate(txtParameter2.Text) < FormatDate(.Text) Then
                        Message "La fecha del campo 2 no puede ser anterior a la del campo 1.", csMsgWarning
                        Exit Sub
                    End If
                End If
              
                Message Msg, Tip
            Exit Sub
        End Select
    End With
End Sub

Private Sub txtParameter2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
       If fAllowSearch = False Then Exit Sub
       Call cmdGo_Click
   End If
End Sub

Private Sub txtParameter2_LostFocus()
    With txtParameter2
        Select Case CurrentFilter
            Case csNumber
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
                
                If Not IsNumeric(.Text) Then
                    Message "El número especificado en el campo 2 no es válido.", csMsgWarning
                    Exit Sub
                End If
                
                If fSearchByTask Then
                    .Text = FormatTaskNumber(Val(.Text))
                Else
                    .Text = FormatReportNumber(Val(.Text))
                End If
                    
                    Message Msg, Tip
            Exit Sub
                      
            Case csDate
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
                
                If Not IsNumeric(.Text) Then
                    Message "La fecha especificada en el campo 2 no es válida.", csMsgWarning
                    Exit Sub
                End If
                
                If Not IsDateValid(.Text) Then
                    Message "La fecha especificada en el campo 2 no corresponde a un día existente.", csMsgWarning
                    Exit Sub
                End If
                    
                    Message Msg, Tip
                Exit Sub
                    
            Case csOkDate
                If Trim(.Text) = "" Then
                    Exit Sub
                End If
                
                If Not IsNumeric(.Text) Then
                    Message "La fecha especificada en el campo 2 no es válida.", csMsgWarning
                    Exit Sub
                End If
                
                If Not IsDateValid(.Text) Then
                    Message "La fecha especificada en el campo 2 no corresponde a un día existente.", csMsgWarning
                    Exit Sub
                End If
                
                If IsDateValid(txtParameter1.Text) Then
                    If FormatDate(.Text) < FormatDate(txtParameter1.Text) Then
                        Message "La fecha del campo 2 no puede ser anterior a la del campo 1.", csMsgWarning
                        Exit Sub
                    End If
                End If
                
                    Message Msg, Tip
                Exit Sub
        End Select
    End With
End Sub

Sub FillParametersComboBox(ByRef cb As ComboBox, ByVal WithWhat As Integer)
    If WithWhat = csPriority Then
        cb.Clear
        cb.AddItem "0"
        cb.AddItem "1"
        cb.AddItem "2"
        cb.AddItem "3"
        
        cb.ListIndex = 1
        
        Exit Sub
    End If
    
    If WithWhat = csStatus Then
        cb.Clear
        cb.AddItem strPEND
        cb.AddItem strOK
        
        cb.ListIndex = 0
        
        Exit Sub
    End If
    
        Dim i As Long
        Dim strTemp As String
    
    ' Corregir índices
        If WithWhat = csBranchOffice Then i = -1
        If WithWhat = csTechnician Then i = -1
        cb.Clear
    
    Do
        i = i + 1
    
        strTemp = ReadIni(Cfg.Data, IIf(WithWhat = csBranchOffice, "SUCURSALES", "TECNICOS"), i, "")
            If strTemp = "" Then Exit Do
        
        cb.AddItem strTemp
    Loop
    
    If cb.ListCount > 0 Then
        cb.ListIndex = 0
    End If

End Sub

Sub FillFilterComboBox(ByVal SearchBy As Integer, ByVal DefaultIndex As Integer)
' Primero cargar datos, despues seleccionar el default
    With cmbSearch
        .Clear
        
        Select Case SearchBy
            Case csByTask
                .AddItem sSearch(csNumber)
                .AddItem sSearch(csDetail)
                .AddItem sSearch(csTechnician)
                .AddItem sSearch(csBranchOffice)
                .AddItem sSearch(csContact)
                .AddItem sSearch(csDate)
                .AddItem sSearch(csOkDate)
                .AddItem sSearch(csStatus)
                .AddItem sSearch(csPriority)
                .AddItem sSearch(csObservations)
                .AddItem sSearch(csHasObservations)
                
            Case csByReport
                .AddItem sSearch(csNumber)
                .AddItem sSearch(csDetail)
                .AddItem sSearch(csTechnician)
                .AddItem sSearch(csBranchOffice)
                .AddItem sSearch(csContact)
                .AddItem sSearch(csDate)
    
        End Select
    
        Select Case DefaultIndex
            Case Is < 0
                If .ListCount > 0 Then .ListIndex = 0
            Exit Sub
              
            Case 0 To 5
                .ListIndex = DefaultIndex
            Exit Sub
                
            Case Is > 5
                .ListIndex = 0
            Exit Sub
        End Select
    End With
End Sub

Sub DisableInput(ByVal Which As Integer, ByVal Status As Boolean)
    Select Case Which
        Case 1
            If Status = True Then
                txtParameter1.BackColor = vbButtonFace
            Else
                txtParameter1.BackColor = vbWindowBackground
            End If
                
            lblParameter1.Enabled = Not Status
            txtParameter1.Enabled = Not Status
            lblParameter1.Caption = ""
            txtParameter1.Text = ""
        Exit Sub
            
        Case 2
            If Status = True Then
                txtParameter2.BackColor = vbButtonFace
            Else
                txtParameter2.BackColor = vbWindowBackground
            End If
            
            lblParameter2.Enabled = Not Status
            txtParameter2.Enabled = Not Status
            lblParameter2.Caption = ""
            txtParameter2.Text = ""
        Exit Sub
    End Select
End Sub

Sub ResetUI(ByVal Mode As Byte)
' Vuelve la interfaz al estado inicial por si se redimensiona por fuera de los límites configurados
    Select Case Mode
        Case cWidth
            Me.Width = 9600
            fraFilter.Move 120, 120
            fraDatos.Move 120, 2040
            fraResults.Move 2760, 120, 6495
            lblDetail.Move 120, lblDetail.Top, 9135
            lstReports.Move 120, lstReports.Top, 6255
            lstTasks.Move 120, lstTasks.Top, 6255
            chkFilterMore.Move 120, chkFilterMore.Top
        Exit Sub
            
        Case cHeight
            Me.Height = 5820
            fraFilter.Move 120, 120
            fraDatos.Move 120, 2040
            fraResults.Move 2760, 120, fraResults.Width, 4695
            lblDetail.Move 120, 4920, lblDetail.Width, 375
            lstReports.Move 120, 300, lstReports.Width, 3960
            lstTasks.Move 120, 300, lstTasks.Width, 3960
            chkFilterMore.Move 120, chkFilterMore.Top
        Exit Sub
        
        Case cBoth
            Me.Move Me.Left, Me.Top, 9600, 5820
            fraFilter.Move 120, 120
            fraDatos.Move 120, 2040
            fraResults.Move 2760, 120, 6495, 4695
            lblDetail.Move 120, 4920, 9135, 375
            lstReports.Move 120, 300, 6255, 3960
            lstTasks.Move 120, 300, 6255, 3960
            chkFilterMore.Move 120, 4320
        Exit Sub
    End Select
End Sub

Sub ShowMessages()
' Hub para mostrar información en base al tipo de búsqueda

    Select Case (cmbSearch.ListIndex + 1)
        Case csNumber
            Msg = "Busca " & sTypes & " en el rango especificado."
            Tip = csMsgInformation
        GoTo ExitCase
            
        Case csDetail
            Msg = "Busca palabras clave dentro de pedidos o informes."
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csTechnician
            Msg = "Busca " & sTypes & " según el técnico asignado."
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csBranchOffice
            Msg = "Busca " & sTypes & " según la sucursal asignada."
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csContact
            Msg = "Busca " & sTypes & " según la persona de contacto."
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csDate
            Msg = "Busca " & sTypes & " en un rango de fechas o una específica. [ddmmyy]"
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csOkDate
            Msg = "Busca pedidos cerrados en un rango de fechas o una específica. [ddmmyy]"
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csStatus
            Msg = "Busca " & sTypes & " según el estado de finalización."
            Tip = csMsgInformation
        GoTo ExitCase
                
        Case csPriority
          Msg = "Filtra pedidos según la prioridad especificada."
            Tip = csMsgInformation
        GoTo ExitCase
        
        Case csObservations
            Msg = "Busca texto en las observaciones de los pedidos."
            Tip = csMsgInformation
        GoTo ExitCase
        
           Case csHasObservations
            Msg = "Muestra sólo los pedidos que tengan observaciones."
            Tip = csMsgInformation
        GoTo ExitCase
    End Select

ExitCase:
    Message Msg, Tip
End Sub

Sub Message(ByVal sMsg As String, ByVal MsgType As Integer)
' VB6 usa color BGR, no RGB
    With lblDetail
        Select Case MsgType
            Case csMsgWarning
                .ForeColor = &H1F1F84
                .Caption = "[!] " & sMsg
            Exit Sub
                
            Case csMsgInformation
                .ForeColor = &H7F3F00
                .Caption = sMsg
            Exit Sub
                
            Case csMsgResult
                .ForeColor = &H7F3F
                .Caption = "[*] " & sMsg
            Exit Sub
                
            Case csMsgNone
                .ForeColor = vbBlack
                .Caption = ""
            Exit Sub
        End Select
    End With
End Sub

Sub AddReport(ByVal ReportsBundle As String)
' Se recibe como |x|x|x
' Filtro es CurrentFilter y se maneja en todo el form

Dim i       As Long     ' Bucle
Dim p       As Long     ' Cada indice del parse

Dim cLen    As Long     ' Longitud máxima de la string de los contactos
Dim cCon()  As String   ' Longitud maxima de la cadena de contactos
Dim cNum()  As String   ' Numeros de pedidos
Dim cDet()  As String   ' Detalles de pedidos

sParse = Split(ReportsBundle, "|")
lstCurrent.Clear

' Personalizar como se despliegan los resultados en base al filtro elegido
Select Case CurrentFilter
    Case csTechnician
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatReportNumber(p) & Space(gSepChar) & _
                        GetTechnician(p, fSearchByTask) & Space(gSepChar) & _
                        Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csBranchOffice
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatReportNumber(p) & Space(gSepChar) & _
                            GetBranchOffice(p, fSearchByTask) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csContact
        ReDim cCon(UBound(sParse))
        ReDim cNum(UBound(sParse))
        ReDim cDet(UBound(sParse))
        
        ' Tomar los datos para no desperdiciar I/O en los dos bucles
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            cNum(i) = FormatReportNumber(p)
            cCon(i) = GetContact(p, fSearchByTask)
            cDet(i) = Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
            
            If Len(cCon(i)) > cLen Then
                cLen = Len(cCon(i))
            End If
        Next i
        
        ' Usar el bucle solo para llenar la lista una vez conocida la maxima longitud (cLen)
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem cNum(i) & Space(gSepChar) & _
                            PadStr(cCon(i), cLen) & Space(gSepChar) & _
                            cDet(i)
        Next
    Exit Sub
    
    Case csDate
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatReportNumber(p) & Space(gSepChar) & _
                            FormatDate2(GetDate(p, fSearchByTask)) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case Else
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatReportNumber(p) & Space(gSepChar) & Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
        Exit Sub
    End Select
End Sub

Sub AddTask(ByVal TaskBundle As String)
' Se recibe como |x|x|x
' Filtro es CurrentFilter y se maneja en todo el form

Dim i       As Long     ' Bucle
Dim p       As Long     ' Cada indice del parse

Dim cLen    As Long     ' Longitud máxima de la string de los contactos
Dim cCon()  As String   ' Longitud maxima de la cadena de contactos
Dim cNum()  As String   ' Numeros de TaskBundle
Dim cDet() As String    ' Detalles de TaskBundle

sParse = Split(TaskBundle, "|")
lstCurrent.Clear

' Personalizar como se despliegan los resultados en base al filtro elegido
Select Case CurrentFilter
    Case csStatus
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            GetStatus(p) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csTechnician
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            GetTechnician(p, fSearchByTask) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csBranchOffice
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            GetBranchOffice(p, fSearchByTask) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csContact
        ReDim cCon(UBound(sParse))
        ReDim cNum(UBound(sParse))
        ReDim cDet(UBound(sParse))
        
        ' Tomar los datos para no desperdiciar I/O en los dos bucles
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            cNum(i) = FormatTaskNumber(p)
            cCon(i) = GetContact(p, fSearchByTask)
            cDet(i) = Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
            
            If Len(cCon(i)) > cLen Then
                cLen = Len(cCon(i))
            End If
        Next i
                
        ' Usar el bucle solo para llenar la lista una vez conocida la maxima longitud (cLen)
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem cNum(i) & Space(gSepChar) & _
                            PadStr(cCon(i), cLen) & Space(gSepChar) & _
                            cDet(i)
        Next
    Exit Sub
    
    Case csDate
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            FormatDate2(GetDate(p, fSearchByTask)) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csOkDate
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            FormatDate2(GetDate(p, fSearchByTask, fSearchByTask)) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case csPriority
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & _
                            GetPriority(p) & Space(gSepChar) & _
                            Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
    Exit Sub
    
    Case Else
        For i = 1 To UBound(sParse)
            p = CLng(sParse(i))
            lstCurrent.AddItem FormatTaskNumber(p) & Space(gSepChar) & Replace(GetDetail(p, fSearchByTask), cNewLine, Space(1))
        Next
        Exit Sub
    End Select
End Sub

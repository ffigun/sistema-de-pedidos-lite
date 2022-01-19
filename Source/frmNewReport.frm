VERSION 5.00
Begin VB.Form frmNewReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTime 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "0900"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtOkDate 
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
      Left            =   4440
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "001122"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtReferenceTask 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox chkOk 
      Caption         =   "Marcar como OK"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtNumber 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "00000"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtContact 
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
      Left            =   2280
      MaxLength       =   23
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox cmbBranchOffice 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtDate 
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
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "001122"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtDetail 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmNewReport.frx":038A
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Label8 
      Caption         =   "Hora:"
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
      TabIndex        =   19
      Top             =   2565
      Width           =   1695
   End
   Begin VB.Image imgValidation 
      Height          =   255
      Left            =   1920
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblOkDate 
      Caption         =   "F. OK:"
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
      Left            =   3720
      TabIndex        =   18
      Top             =   2085
      Width           =   615
   End
   Begin VB.Label lblDetailLen 
      Alignment       =   1  'Right Justify
      Caption         =   "255 / 255"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Pedido de Ref:"
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
      TabIndex        =   16
      Top             =   645
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Número:"
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
      TabIndex        =   15
      Top             =   165
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Contacto:"
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
      TabIndex        =   14
      Top             =   1125
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Sucursal:"
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
      TabIndex        =   13
      Top             =   1605
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha:"
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
      Top             =   2085
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Detalle:"
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
      Top             =   3045
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fCtrlReturn As Boolean

Private Sub chkOk_Click()
    DisplayOkDateTextBox (chkOk.Value)
End Sub

Private Sub cmdSave_Click()
Dim Reason              As String
Dim MsgString           As String       ' Cadena del msgbox
Dim OverrideTechnician  As Boolean      ' Sobreescribir el técnico actual en caso de editar el informe

' Validar
    If Not fNewReport Then
        If Not ReportExists(CLng(txtNumber.Text)) Then
            Reason = "El informe especificado no existe."
            GoTo OhNo
        End If
    End If

    If Not TaskExists(CLng(txtReferenceTask.Text)) Then
        Reason = "El pedido especificado no existe."
        GoTo OhNo
    End If
    
    If Len(Trim(txtContact.Text)) = 0 Then
        Reason = "Debe especificar un contacto."
        GoTo OhNo
    End If
    
    If Not IsNumeric(txtDate.Text) Then
        Reason = "Sólo puede especificar números en los campos de fecha, en el formato ddmmyy."
        GoTo OhNo
    End If
    
    If Not IsDateValid(txtDate.Text) Then
        Reason = "Debe indicar una fecha válida."
        GoTo OhNo
    End If
    
    If fNewReport Then
        If GetStatus(CLng(txtReferenceTask.Text)) = strOK Then
            Reason = "El pedido existe pero está cerrado (Estado OK). Para volver a abrirlo ingrese al menú de pedidos, edítelo y cámbielo a PEND quitándo la marca OK."
            GoTo OhNo
        End If
    End If
    
    If chkOk.Value = vbChecked Then
        If Not IsNumeric(txtOkDate.Text) Then
            Reason = "Sólo puede especificar números en los campos de fecha, en el formato ddmmyy."
            GoTo OhNo
        End If
        
        If Not IsDateValid(txtOkDate.Text) Then
            Reason = "Debe indicar una fecha válida."
            GoTo OhNo
        End If
    End If
    
    If Len(Trim(txtDetail.Text)) = 0 Or Len(Trim(Replace(txtDetail.Text, vbNewLine, ""))) = 0 Then
        Reason = "Debe especificar un detalle."
        GoTo OhNo
    End If
    
    If Not IsTimeValid(txtTime.Text) Then
        Reason = "Debe indicar una hora válida."
        GoTo OhNo
    End If

' Capitalizar
    Call txtDetail_LostFocus
    Call txtContact_LostFocus
    
' Si esta editando, verificar si quiere pisar el técnico que cargó el informe
    If Not fNewReport Then
        If GetTechnician(Report.Number, False) <> CurrentTechnician Then
            MsgString = "El informe que está editando fue cargado por " & GetTechnician(Report.Number, False) & "." & vbNewLine & vbNewLine & _
                        "¿Desea sobreescribirlo por " & CurrentTechnician & "?"
                          
            Select Case MsgBox(MsgString, vbQuestion + vbYesNoCancel, "Sobreescribir técnico")
                Case vbYes:      OverrideTechnician = True
                Case vbNo:       OverrideTechnician = False
                Case vbCancel:  Exit Sub
            End Select
        End If
    End If

' Asignar un numero sólo si es un nuevo informe. Ademas, hacer que el tecnico actual sea el que se grabe en el informe
    If fNewReport Then
        txtNumber.Text = FormatReportNumber(FreeReport)
        OverrideTechnician = True
        
        If fDebug Then
            Log "frmNewReport asked for a free report number. It was given the number " & FreeReport & " which formatted is " & FormatReportNumber(FreeReport)
        End If
    End If
    
    With TempReport
        .Contact = txtContact.Text
        .Detail = txtDetail.Text
        .Date = txtDate.Text
        .Time = txtTime.Text
        .Number = CLng(txtNumber.Text)
        .ReferenceTask = CLng(txtReferenceTask.Text)
        .BranchOffice = cmbBranchOffice.List(cmbBranchOffice.ListIndex)
        If OverrideTechnician Then
            .Technician = CurrentTechnician
        Else
            .Technician = GetTechnician(CurrentReport, False)
        End If

        If chkOk.Value = vbChecked Then
            CloseTask .ReferenceTask, .Number, txtOkDate.Text
            
            If fDebug Then
                Log "The procedure CloseTask was invoked with parameters " & .ReferenceTask & " , " & .Number & " , " & txtOkDate.Text
            End If

        End If
    End With

' Asignar el informe actual a CurrentReport para que al llamar a .Update del visor vaya al ultimo!
    CurrentReport = TempReport.Number
    WriteReport TempReport
    
    If fNewReport Then AddOneToLastReport
        fAlreadySetFocus = False
            If chkOk.Value = vbChecked Or fHlTasksWithReports = True Then
                frmMain.AddTask (MainFilter)
            End If
    
    frmMain.Update
    
    If fDebug Then
        Log "Leaving frmNewReport, NewReport flag is " & fNewReport & vbNewLine & _
            "Current Report is " & CurrentReport & vbNewLine & _
            "Current Task is " & TaskHolder
    End If

' Si se llamó al modo edición desde el visor de informes o el visor de pedidos
    If fFromFrmReportViewer = True Then
        frmReportViewer.Update
    End If
        
    If fFromFrmTaskViewer = True Then
        frmTaskViewer.Update
    End If
    
    Unload Me
Exit Sub

OhNo:
    MsgBox Reason, vbExclamation, "Compruebe los datos ingresados y vuelva a intentarlo."
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    fAlreadySetFocus = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Si aprieta Ctrl + Enter
    If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
        Call cmdSave_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' Deshacerse de los caracteres 127 (Ctrl + Backspace) y el 10 (Ctrl + Enter)
    If KeyAscii = 10 Then KeyAscii = 0
    If KeyAscii = 127 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim i As Long

    If fDebug Then
        Log "Entering frmNewReport, NewReport flag is " & fNewReport & vbNewLine & _
            "Current Report is " & CurrentReport & vbNewLine & _
            "Current Task is " & TaskHolder
    End If
    
    If fNewReport Then
    ' Cargar nuevo informe
        Me.Caption = "Nuevo informe"
         
        LoadBranchOffices cmbBranchOffice
        chkOk.Value = vbUnchecked
        DisplayOkDateTextBox (False)
         
        txtNumber.Text = "" ' No asignar hasta guardar
        txtReferenceTask.Text = IIf(TaskHolder < 0, "", FormatTaskNumber(TaskHolder))
        txtContact.Text = IIf(TaskHolder < 0, "", GetContact(TaskHolder, True))
        txtContact.SelStart = 0
        txtContact.SelLength = Len(txtContact.Text)
        txtDate.Text = PlainDate(Date)
        txtOkDate.Text = PlainDate(Date)
        txtTime.Text = PlainTime(Time)
        txtDetail.Text = ""
        DisplayLen txtDetail, lblDetailLen
        cmbBranchOffice.ListIndex = FindItemInComboBox(cmbBranchOffice, GetBranchOffice(TaskHolder, True))
    Else
    ' Editar informe
        Report = ReadReport(CurrentReport)
        
        With Report
           Me.Caption = "Editar informe " & FormatReportNumber(.Number)
           
           LoadBranchOffices cmbBranchOffice
           SetIndexByText cmbBranchOffice, .BranchOffice
           chkOk.Value = IIf(GetStatus(.ReferenceTask) = strOK, vbChecked, vbUnchecked)
           DisplayOkDateTextBox (chkOk.Value)
           
           txtNumber.Text = FormatReportNumber(.Number)
           txtReferenceTask.Text = FormatTaskNumber(.ReferenceTask)
           txtContact.Text = .Contact
           txtDate.Text = .Date
           txtOkDate.Text = GetOkDate(.ReferenceTask)
           txtTime.Text = .Time
           txtDetail.Text = .Detail
           cmbBranchOffice.ListIndex = FindItemInComboBox(cmbBranchOffice, .BranchOffice)
           
           DisplayLen txtDetail, lblDetailLen
        End With
        
        ' Si el pedido está OK
        If Trim(txtReferenceTask.Text) <> "" Then
           If TaskExists(CLng(txtReferenceTask.Text)) Then
               If UCase(GetStatus(CLng(txtReferenceTask.Text))) = strOK Then
                   chkOk.Value = vbChecked
                   chkOk.Enabled = False
               Else
                   chkOk.Value = vbUnchecked
                   chkOk.Enabled = True
               End If
           End If
        End If
    End If
End Sub

Private Sub Form_Paint()
    If fAlreadySetFocus Then Exit Sub

    If TaskHolder >= 0 Then
        txtContact.SetFocus
        fAlreadySetFocus = True
    End If
End Sub

Private Sub txtDetail_Change()
    DisplayLen txtDetail, lblDetailLen
End Sub

Private Sub txtDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
        DelLastWord txtDetail
    End If
End Sub

Private Sub txtContact_GotFocus()
    txtContact.SelStart = 0
    txtContact.SelLength = Len(txtContact.Text)
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
        DelLastWord txtContact
    End If
End Sub

Private Sub txtContact_LostFocus()
    txtContact.Text = StrConv(txtContact.Text, vbProperCase)
End Sub

Private Sub txtDetail_GotFocus()
    txtDetail.SelStart = Len(txtDetail.Text)
End Sub

Private Sub txtDetail_LostFocus()
On Error Resume Next
    With txtDetail
         .Text = UCase(Mid(.Text, 1, 1)) & Right(.Text, Len(.Text) - 1)
    End With
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_LostFocus()
    With imgValidation
         .Visible = False
         .Move txtDate.Left - .Width - 60, txtDate.Top + 60
         
        If Len(Trim(txtDate.Text)) = 0 Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Debe ingresar una fecha."
           .Visible = True
           Exit Sub
        End If
         
        If Not IsNumeric(txtDate.Text) Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Ingrese sólo números en el formato ddmmyy."
           .Visible = True
           Exit Sub
        Else
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La fecha es válida."
           .Visible = True
        End If
         
        If IsDateValid(txtDate.Text) Then
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La fecha es válida."
           .Visible = True
        Else
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "La fecha no es válida."
           .Visible = True
           Exit Sub
        End If
    End With
End Sub

Private Sub txtOkDate_GotFocus()
    txtOkDate.SelStart = 0
    txtOkDate.SelLength = Len(txtOkDate.Text)
End Sub

Private Sub txtOkDate_LostFocus()
    With imgValidation
         .Visible = False
         .Move lblOkDate.Left - .Width - 60, lblOkDate.Top + 60
         
        If Len(Trim(txtOkDate.Text)) = 0 Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Debe ingresar una fecha."
           .Visible = True
           Exit Sub
        End If
         
        If Not IsNumeric(txtOkDate.Text) Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Ingrese sólo números en el formato ddmmyy."
           .Visible = True
           Exit Sub
        Else
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La fecha es válida."
           .Visible = True
        End If
         
        If IsDateValid(txtOkDate.Text) Then
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La fecha es válida."
           .Visible = True
        Else
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "La fecha no es válida."
           .Visible = True
           Exit Sub
        End If
    End With
End Sub

Private Sub txtTime_GotFocus()
    txtTime.SelStart = 0
    txtTime.SelLength = Len(txtTime.Text)
End Sub

Private Sub txtTime_LostFocus()
    With imgValidation
         .Visible = False
         .Move txtTime.Left - .Width - 60, txtTime.Top + 60
         
        If Len(Trim(txtTime.Text)) = 0 Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Debe ingresar una hora."
           .Visible = True
           Exit Sub
        End If
         
        If Not IsNumeric(txtTime.Text) Then
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "Ingrese sólo números en el formato hhmm."
           .Visible = True
           Exit Sub
        Else
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La hora es válida."
           .Visible = True
        End If
         
        If IsTimeValid(txtTime.Text) Then
           .Picture = frmMain.imgCheck.Picture
           .ToolTipText = "La hora es válida."
           .Visible = True
        Else
           .Picture = frmMain.imgError.Picture
           .ToolTipText = "La hora no es válida."
           .Visible = True
           Exit Sub
        End If
         .Visible = True
    End With
End Sub

Sub SetIndexByText(ByRef cb As ComboBox, ByVal sBranchOffice As String)
Dim i As Long

    For i = 0 To cb.ListCount
        If Trim(LCase(cb.List(i))) = Trim(LCase(sBranchOffice)) Then
           cb.ListIndex = i
           Exit For
        End If
    Next
End Sub

Sub DisplayLen(ByRef tb As TextBox, ByRef lb As Label)
    lb.Caption = Len(tb.Text) & " / " & tb.MaxLength
End Sub

Sub DisplayOkDateTextBox(ByVal bCurrentValue As Boolean)
    txtOkDate.Enabled = bCurrentValue
    txtOkDate.Visible = bCurrentValue
     
    lblOkDate.Enabled = bCurrentValue
    lblOkDate.Visible = bCurrentValue
End Sub

Sub LoadBranchOffices(ByRef cb As ComboBox)
Dim i As Long: i = -1
Dim strTemp As String
    
    cb.Clear

Do
    i = i + 1
    strTemp = ReadIni(Cfg.Data, "SUCURSALES", i, "")
        
        If strTemp = "" Then
            Exit Do
        End If
        
    cb.AddItem strTemp
Loop

    If cb.ListCount > 0 Then
        cb.ListIndex = 0
    End If

End Sub

VERSION 5.00
Begin VB.Form frmNewTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo pedido"
   ClientHeight    =   6450
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
   Icon            =   "frmNewTask.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
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
      TabIndex        =   5
      Text            =   "0900"
      Top             =   2040
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
      TabIndex        =   4
      Text            =   "001122"
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkOk 
      Caption         =   "El pedido está OK"
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
      Height          =   375
      Left            =   3285
      TabIndex        =   0
      Top             =   120
      Width           =   2130
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
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
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
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtObservations 
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
      MaxLength       =   64
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   5400
      Width           =   5175
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
      TabIndex        =   7
      Text            =   "frmNewTask.frx":038A
      Top             =   3480
      Width           =   5175
   End
   Begin VB.OptionButton optPriority 
      Caption         =   "0"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   2520
      Width           =   495
   End
   Begin VB.OptionButton optPriority 
      Caption         =   "3"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.OptionButton optPriority 
      Caption         =   "2"
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
      Left            =   3480
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.OptionButton optPriority 
      Caption         =   "1"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   495
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
      TabIndex        =   3
      Text            =   "001122"
      Top             =   1560
      Width           =   855
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
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
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
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   3015
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
      TabIndex        =   11
      Text            =   "00000"
      Top             =   120
      Width           =   855
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
      TabIndex        =   25
      Top             =   2085
      Width           =   1695
   End
   Begin VB.Image imgValidation 
      Height          =   255
      Left            =   1920
      Top             =   1560
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
      TabIndex        =   24
      Top             =   1605
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblObservationsLen 
      Alignment       =   1  'Right Justify
      Caption         =   "64 / 64"
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
      TabIndex        =   23
      Top             =   5040
      Width           =   1095
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
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Observaciones:"
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
      TabIndex        =   21
      Top             =   4965
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
      TabIndex        =   20
      Top             =   3045
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Prioridad:"
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
      Top             =   2565
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
      TabIndex        =   15
      Top             =   1605
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
      TabIndex        =   14
      Top             =   1125
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   165
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOk_Click()
    DisplayOkDateTextBox chkOk.Value
End Sub

Private Sub cmdSave_Click()
Dim Reason                  As String
Dim MsgString               As String
Dim OverrideTechnician      As Boolean

' Validar
    If Not fNewTask Then
        If Not TaskExists(CLng(txtNumber.Text)) Then
           Reason = "El pedido especificado no existe."
           GoTo OhNo
        End If
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
    Call txtObservations_LostFocus
    
' Si esta editando, verificar si quiere pisar el técnico que cargó el informe
    If Not fNewTask Then
       If GetTechnician(Task.Number, True) <> CurrentTechnician Then
          MsgString = "El pedido que está editando fue cargado por " & GetTechnician(Task.Number, True) & "." & vbNewLine & vbNewLine & _
                      "¿Desea sobreescribirlo por " & CurrentTechnician & "?"
                           
          Select Case MsgBox(MsgString, vbQuestion + vbYesNoCancel, "Sobreescribir técnico")
                Case vbYes:      OverrideTechnician = True
                Case vbNo:       OverrideTechnician = False
                Case vbCancel:  Exit Sub
          End Select
       End If
    End If
    
' Asignar numero de pedido sólo si es un nuevo pedido. Ademas, hacer que el tecnico actual sea el que se grabe en el informe
    If fNewTask Then
        txtNumber.Text = FormatTaskNumber(FreeTask)
        OverrideTechnician = True
         
        If fDebug Then
           Log "frmNewTask asked for a free task number. It was given the number " & FreeTask & " which formatted is " & FormatTaskNumber(FreeTask)
        End If
    End If
     
    With TempTask
        .Contact = txtContact.Text
        .Detail = txtDetail.Text
        .Status = IIf(chkOk.Value, strOK, strPEND)
        .OkDate = IIf(chkOk.Value = vbChecked, txtOkDate.Text, "")
        .Date = txtDate.Text
        .Time = txtTime.Text
        .Number = txtNumber.Text
        .Observations = txtObservations.Text
        .BranchOffice = cmbBranchOffice.List(cmbBranchOffice.ListIndex)
        If OverrideTechnician Then
            .Technician = CurrentTechnician
        Else
            .Technician = GetTechnician(TaskHolder, True)
        End If
        .OkReport = IIf(chkOk.Value = vbChecked, Task.OkReport, -1)
     
        If optPriority(0).Value = True Then .Priority = 0
        If optPriority(1).Value = True Then .Priority = 1
        If optPriority(2).Value = True Then .Priority = 2
        If optPriority(3).Value = True Then .Priority = 3
    End With

    WriteTask TempTask

    If fNewTask Then AddOneToLastTask
    
    If Not fNewTask Then
        If chkOk.Value = vbChecked Then
            CloseTask CLng(txtNumber.Text), GetOkReport(CLng(txtNumber.Text)), txtOkDate.Text
          
            If fDebug Then
                Log "frmNewTask asked for a free task number. It was given the number " & FreeTask & " which formatted is " & FormatTaskNumber(FreeTask)
            End If
        Else
            OpenClosedTask (CLng(txtNumber.Text))
          
            If fDebug Then
                Log "The procedure OpenClosedTask was invoked with parameter " & (CLng(txtNumber.Text)) & "."
            End If
        End If
    End If

' Si cargo un nuevo informe, levantar el flag para que al actualizar la lista de pedidos siempre muestre el ultimo
    If fNewTask Then
        fFromFrmNewTask = True
    End If
    
    frmMain.AddTask (MainFilter)
    frmMain.Update
    
' Si se llamó al modo edición desde el visor de pedidos
   If fFromFrmTaskViewer Then
       frmTaskViewer.Update
   End If

' Si estaba abierta la ventana del visor de informes, actualizar tambien
   If fFromFrmReportViewer Then
       frmReportViewer.Update
   End If
    
   Unload Me
    
Exit Sub
    
OhNo:
    MsgBox Reason, vbExclamation, "Compruebe los datos ingresados y vuelva a intentarlo."
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        Call cmdSave_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' Deshacerse de los caracteres 127 (Ctrl + Backspace) y el 10 (Ctrl + Enter)
    If KeyAscii = 10 Then KeyAscii = 0
    If KeyAscii = 127 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    'ApplyFont txtDetail, fViewers
    'ApplyFont txtObservations, fViewers
    
    If fNewTask Then
    ' Cargar nuevo pedido
        Me.Caption = "Nuevo pedido"
        txtNumber = ""  ' No asignar numero hasta guardarlo
        txtContact.Text = ""
        LoadBranchOffices cmbBranchOffice
        txtDate.Text = PlainDate(Date)
        txtOkDate.Text = PlainDate(Date)
        txtTime.Text = PlainTime(Time)
        chkOk.Enabled = False
        chkOk.Visible = False
        DisplayOkDateTextBox (False)
        optPriority(1).Value = True
        txtDetail.Text = ""
        txtObservations.Text = ""
         
        DisplayLen txtDetail, lblDetailLen
        DisplayLen txtObservations, lblObservationsLen
    Else
    ' Editar pedido
        Task = ReadTask(TaskHolder)
    
        With Task
            Me.Caption = "Editar pedido " & FormatTaskNumber(.Number)
           
            If .Status = strOK Then
                chkOk.Enabled = True
                chkOk.Value = vbChecked
            Else
                chkOk.Enabled = False
                chkOk.Value = vbUnchecked
            End If
                
            txtNumber = FormatTaskNumber(.Number)
            txtContact.Text = .Contact
            LoadBranchOffices cmbBranchOffice
            SetIndexByText cmbBranchOffice, .BranchOffice
            txtDate.Text = .Date
            txtOkDate.Text = .OkDate
            txtTime.Text = .Time
            DisplayOkDateTextBox (chkOk.Value)
            optPriority(.Priority).Value = True
            
            txtDetail.Text = .Detail
            txtObservations.Text = .Observations
            
            DisplayLen txtDetail, lblDetailLen
            DisplayLen txtObservations, lblObservationsLen
        End With
    End If
End Sub

Private Sub txtContact_LostFocus()
    txtContact.Text = StrConv(txtContact.Text, vbProperCase)
End Sub

Private Sub txtDetail_Change()
    DisplayLen txtDetail, lblDetailLen
End Sub

Private Sub txtDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
        DelLastWord Me.txtDetail
    End If
End Sub

Private Sub txtDetail_LostFocus()
On Error Resume Next
    With txtDetail
         .Text = UCase(Mid(.Text, 1, 1)) & Right(.Text, Len(.Text) - 1)
    End With
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
        .Visible = True
    End With
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
        .Visible = True
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

Private Sub txtObservations_Change()
    DisplayLen txtObservations, lblObservationsLen
End Sub

Private Sub txtObservations_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
        DelLastWord txtObservations
    End If
End Sub

Private Sub txtObservations_LostFocus()
On Error Resume Next
    With txtObservations
        .Text = UCase(Mid(.Text, 1, 1)) & Right(.Text, Len(.Text) - 1)
    End With
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
' Si aprieta Ctrl + Enter
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Call cmdSave_Click
End Sub

Private Sub txtContact_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack And Shift = vbCtrlMask Then
        DelLastWord txtContact
    End If
End Sub

Private Sub txtContact_GotFocus()
    txtContact.SelStart = 0
    txtContact.SelLength = Len(txtContact.Text)
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtOkDate_GotFocus()
    txtOkDate.SelStart = 0
    txtOkDate.SelLength = Len(txtOkDate.Text)
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

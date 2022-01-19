VERSION 5.00
Begin VB.Form frmBeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funciones experimentales"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstBeta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Ejecutar"
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
      TabIndex        =   0
      Top             =   1725
      Width           =   4575
   End
End
Attribute VB_Name = "frmBeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
Dim sBuffer As String

With lstBeta
    If .ListIndex = -1 Then Exit Sub
    
    Select Case Left(.List(.ListIndex), 1)
        Case "1"
            frmDebug.Show vbModeless
            Exit Sub
            
        Case "2"
            Call ConvertirInisAV2
            Exit Sub
            
        Case "3"
            fDebug = Not fDebug
            lstBeta.List(2) = "3 Activar variable fDebug (Actual: " & fDebug & ")"
            Exit Sub
            
        Case "4"
            lstBeta.Enabled = False
            cmdGo.Enabled = False
            lstBeta.List(3) = "4 Estad�sticas (Calculando...)"
            Screen.MousePointer = vbHourglass
            
            sBuffer = "Cantidades:" & vbNewLine & _
                      "Pedidos: " & LastTaskFull & vbNewLine & _
                      "Informes: " & LastReportFull & vbNewLine & vbNewLine & _
                      "Tama�o de archivos:" & vbNewLine & _
                      "Pedidos: " & Format(FileLen(Cfg.Tasks), "##,##") & " Bytes" & vbNewLine & _
                      "Informes: " & Format(FileLen(Cfg.Reports), "##,##") & "Bytes" & vbNewLine & _
                      "Enlace: " & Format(FileLen(Cfg.Link), "##,##") & " Bytes" & vbNewLine & vbNewLine & _
                      "�ltimos �ndices:" & vbNewLine & _
                      "Pedidos: " & LastTask & vbNewLine & _
                      "Informes: " & LastReport
            
            lstBeta.Enabled = True
            cmdGo.Enabled = True
            lstBeta.List(3) = "4 Estad�sticas"
            Screen.MousePointer = vbNormal
            
            MsgBox sBuffer, vbInformation + vbOKOnly, "Estad�sticas"
            Exit Sub
    End Select
End With
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    Call LlenarLista
End Sub

Sub LlenarLista()
    With lstBeta
        .AddItem "1 Mostrar �ndices de TaskHolder"
        .AddItem "2 Convertir INIs de V1 a V2"
        .AddItem "3 Activar variable fDebug (Actual: " & fDebug & ")"
        .AddItem "4 Estad�sticas"
    End With
End Sub

Sub ConvertirInisAV2()
If MsgBox("Esta es una funci�n experimental que puede destruir por completo los archivos de Pedidos e Informes." & vbNewLine & vbNewLine & _
          "�Est� seguro que desea continuar?", vbCritical + vbYesNo, "Funci�n de transici�n") = vbNo Then Exit Sub
            
MsgBox "Deber� realizar los siguientes pasos:" & vbNewLine & vbNewLine & _
       "1. Copie los archivos INI de Informes y Pedidos de la versi�n 1 a la carpeta ra�z del programa." & vbNewLine & _
       "2. Ren�mbrelos a INFORMESV1.INI y PEDIDOSV1.INI respetando las may�sculas." & vbNewLine & vbNewLine & _
       "Si ya realiz� este procedimiento, presione S�. De lo contrario, presione No y real�celo.", vbInformation + vbYesNoCancel, "Procedimiento previo"
       
Dim a As Long
Dim b As Long
Dim s As Boolean
Dim INI_Inf As String: INI_Inf = App.Path & "\INFORMESV1.INI"
Dim INI_Ped As String: INI_Ped = App.Path & "\PEDIDOSV1.INI"

' Existencia
    If Dir$(INI_Inf) = "" Then
        MsgBox "No se encontr� el archivo �INFORMESV1.INI� en la ra�z del programa. El proceso no puede continuar.", vbExclamation
        Exit Sub
    End If
    
    If Dir$(INI_Ped) = "" Then
        MsgBox "No se encontr� el archivo �PEDIDOSV1.INI� en la ra�z del programa. El proceso no puede continuar.", vbExclamation
        Exit Sub
    End If
    
' Sigamos
a = 0
b = 0
s = False

' Pedidos
    Do Until s = True
        If ReadIni(INI_Ped, Format(Str(a), "00000"), "DE", "") <> "" Then
            a = a + 1
        Else
            s = True
        End If
    Loop
    
' Informes
s = False
    
    Do Until s = True
        If ReadIni(INI_Inf, Format(Str(b), "00000"), "DE", "") <> "" Then
            b = b + 1
        Else
            s = True
        End If
    Loop

' Ultima advertencia
If MsgBox("Se encontraron " & a & " pedido(s) y " & b & " informe(s)." & vbNewLine & vbNewLine & _
          "�Desea crear dos copias V2?", vbQuestion + vbYesNo, "Confirmar creaci�n de copias V2") = vbNo Then Exit Sub

    MsgBox "El proceso puede demorar unos minutos. No interrumpa la operaci�n.", vbExclamation + vbOKOnly
        
    cmdGo.Enabled = False
    lstBeta.Enabled = False
    Screen.MousePointer = vbHourglass
        
Dim strBuffer As String
Dim strNumber As String
Dim f As Long

' Pedidos (-1 porque incluye el cero)
    For f = 0 To a - 1
        strNumber = Format(Str(f), "00000")
    
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "CT") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "DE") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "ST") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "OD") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "ID") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "TI") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "OB") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "PR") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "BO") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "TN") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Ped, strNumber, "OR")
        
        WriteIni App.Path & "\TasksV2Converted.ini", strNumber, cData, strBuffer
        lstBeta.List(1) = "2 Convertir INIs de V1 a V2 (" & Round(f * 100 / (a + b), 0) & "%)"
        strBuffer = ""
        
        DoEvents
    Next
    
' Informes (-1 porque incluye el cero)
    For f = 0 To b - 1
        strNumber = Format(Str(f), "00000")
    
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "CT") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "DE") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "ID") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "TI") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "BO") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "TN") & cSepChar
        strBuffer = strBuffer & ReadIni(INI_Inf, strNumber, "RT")
        
        WriteIni App.Path & "\ReportsV2Converted.ini", strNumber, cData, strBuffer
        lstBeta.List(1) = "2 Convertir INIs de V1 a V2 (" & Round((f + a) * 100 / (a + b), 0) & "%)"
        strBuffer = ""
        
        DoEvents
    Next
    
    MsgBox "Los archivos se generaron en:" & vbNewLine & App.Path & vbNewLine & vbNewLine & _
           "Los mismos se llaman �TasksV2Converted.ini� y �ReportsV2Converted.ini�.", vbInformation
           
    MsgBox "El proceso se complet�." & vbNewLine & vbNewLine & "Copie el archivo �Enlace.ini� de la carpeta del progama del que convirti� los archivos, de lo contrario los informes estar�n desordenados, se da�ar�n los �ndices y se podr�a sobreescribir permanentemente la informaci�n.", vbInformation
    
    lstBeta.List(1) = "2 Convertir INIs de V1 a V2"
    
    cmdGo.Enabled = True
    lstBeta.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub lstBeta_DblClick()
    Call cmdGo_Click
End Sub

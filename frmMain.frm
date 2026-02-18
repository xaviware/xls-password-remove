VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Eliminator .XLS"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4530
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStatus 
      Interval        =   1000
      Left            =   4800
      Top             =   315
   End
   Begin VB.TextBox txtResult 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1440
      Width           =   8055
   End
   Begin MSComDlg.CommonDialog OpenFile 
      Left            =   4200
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Xls"
      DialogTitle     =   "Choose Excel File to Find Password"
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   5400
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C91C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F96E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B9C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBuyNow 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      Picture         =   "frmMain.frx":1C474
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1440
      Picture         =   "frmMain.frx":1C7BE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdObtener 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      Picture         =   "frmMain.frx":1D088
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":1D952
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label lblResultado 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   8055
   End
   Begin VB.Label lblArchivo 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblLnkWeb 
      Caption         =   "http://www.xaviware.cl"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6360
      MouseIcon       =   "frmMain.frx":1DD51
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblRegistrado 
      Caption         =   "Software Registrado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label lblTimer 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuIdioma 
         Caption         =   "&Idioma"
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuLang 
            Caption         =   ""
            Index           =   9
         End
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "Archivos Recientes"
         Begin VB.Menu mnuReciente 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuReciente 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuReciente 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuReciente 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuReciente 
            Caption         =   ""
            Index           =   4
         End
      End
      Begin VB.Menu mnuProducts 
         Caption         =   "Otros Productos"
         Begin VB.Menu mnuWebbrowsers 
            Caption         =   "Web Browsers"
            Begin VB.Menu mnuPWDIE 
               Caption         =   "Internet Explorer Password Recovery Utility"
            End
            Begin VB.Menu mnuPWDMFF 
               Caption         =   "Mozilla Firefox Password Recovery"
            End
            Begin VB.Menu mnuPWDCHR 
               Caption         =   "Google Chrome Password Recovery"
            End
            Begin VB.Menu mnuCAPE 
               Caption         =   "Content Advisor Password Eliminator"
            End
            Begin VB.Menu mnuPWDSafari 
               Caption         =   "Safari Password Recovery"
            End
            Begin VB.Menu mnulin0 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPWDFlock 
               Caption         =   "Flock Password Recovery"
            End
            Begin VB.Menu mnuPWDChromePlus 
               Caption         =   "ChromePlus Password Recovery"
            End
            Begin VB.Menu mnuPWDRAGON 
               Caption         =   "Comodo Dragon Password Recovery"
            End
            Begin VB.Menu mnuRockPWD 
               Caption         =   "RockMelt Password Recovery"
            End
            Begin VB.Menu mnuPWDIRON 
               Caption         =   "SRWare Iron Password Recovery"
            End
         End
         Begin VB.Menu mnuG1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPWDMSN 
            Caption         =   "Windows Live Messenger Password Recovery"
         End
         Begin VB.Menu mnuPWDGTALK 
            Caption         =   "Google Talk Password Decryptor"
         End
         Begin VB.Menu mnuPWDTRILLIAN 
            Caption         =   "Trillian and Trillian Astra Password Recovery"
         End
         Begin VB.Menu mnuPIDGIN 
            Caption         =   "Pidgin Password Recovery"
         End
         Begin VB.Menu mnuG2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPWDFB 
            Caption         =   "Facebook Password Recovery"
         End
         Begin VB.Menu mnuPWDTW 
            Caption         =   "Twitter Password Recovery"
         End
         Begin VB.Menu mnuPWDGMAIL 
            Caption         =   "Google Mail Password Recovery"
         End
         Begin VB.Menu mnuG4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPWDMDB 
            Caption         =   "Password Recovery .MDB"
         End
         Begin VB.Menu mnuRRPDF 
            Caption         =   "Restrictions Remover .PDF"
         End
         Begin VB.Menu mnuG3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPWDFZ 
            Caption         =   "FileZilla Password Recovery Tool"
         End
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyudaF1 
         Caption         =   "Ayuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "A&cerca de ..."
      End
      Begin VB.Menu mnulin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompra 
         Caption         =   "Compre en Línea"
      End
      Begin VB.Menu mnuRegistro 
         Caption         =   "Registrar..."
      End
      Begin VB.Menu mnulin2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFeedBack 
         Caption         =   "Enviar Sugerencia ó Reclamo"
      End
      Begin VB.Menu mnuGoToWeb 
         Caption         =   "Ir a la Web de Password Eliminator .XLS"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : frmMain
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Formulario Principal
'---------------------------------------------------------------------------------------

Private Sub cmdOpenFile_Click()
10        On Error Resume Next
20        OpenFile.FileName = ""
30        OpenFile.Filter = "Excel Files (*.Xls)|*.Xls"
40        OpenFile.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
50        OpenFile.ShowOpen
60        Me.lblFile.Caption = OpenFile.FileName
70        Me.txtResult.Text = Empty
End Sub

Private Sub cmdRecoverPwd_Click()
'---------------------------------------------------------------------------------------
' Procedure : cmdRecoverPwd_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Obtiene Contraseñas
'---------------------------------------------------------------------------------------
10       On Error GoTo cmdRecoverPwd_Click_Error

20        If Len(Me.lblFile.Caption) = 0 Then
30            MsgBox "No ha Seleccionado un Archivo" & vbCrLf & "Seleccione un Archivo Primero...", vbCritical, "Password Recovery .XLS"
40            Exit Sub
50        End If
              
60        Set ExcelApp = CreateObject("Excel.Application")
70        strWorkBookPwd = Empty
          
80        Me.txtResult.Text = Empty
90        Screen.MousePointer = vbHourglass
100       Call GetAllPasswords
110       Screen.MousePointer = vbDefault
120       Me.tmrStatus.Enabled = False

130      On Error GoTo 0
140      Exit Sub

cmdRecoverPwd_Click_Error:
150       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRecoverPwd_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Public Sub cmdBuyNow_Click()
'---------------------------------------------------------------------------------------
' Procedure : cmdBuyNow_Click
' Author    : XaviWare
' Date      : 30-05-2010
' Purpose   : Botón Comprar
'---------------------------------------------------------------------------------------
'
    Call Me.mnuCompra_Click
End Sub

Private Sub cmdObtener_Click()
    Call Me.btnObtener
End Sub

Private Sub cmdOpen_Click()
    Call Me.btnExaminar
End Sub

Private Sub cmdSalir_Click()
'---------------------------------------------------------------------------------------
' Procedure : cmdSalir_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Finalizo el Programa
'---------------------------------------------------------------------------------------
10       On Error GoTo cmdSalir_Click_Error

15       Unload Me
20       End

30       On Error GoTo 0
40       Exit Sub

cmdSalir_Click_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

' ------------------------------------------------------------------------------------------------------------
' Funciones Principales
' ------------------------------------------------------------------------------------------------------------

Private Function GetAllPasswords()
'---------------------------------------------------------------------------------------
' Procedure : GetAllPasswords
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Función de Proceso Principal de Recuperación
'---------------------------------------------------------------------------------------
10       On Error GoTo GetAllPasswords_Error

20        dteInicio = Now
30        lngPwdCount = 0
40        Me.tmrStatus.Enabled = True
          
          'Verifico si el Archivo Tiene Contraseña
50        Call Add2Status(MSGINITIALIZE & vbNewLine)
60        If CheckPwdFileSet Then
              'Busco la Contraseña del Archivo
70            If Not blnCantOpen Then
80                If Not modGetPwd.BasedOnValidPwdsAttackWorkbook Then
90                   Call BruteForceAttackWorkBook
100               End If
110           End If
120       Else
130           Call Add2Status(Replace(MSGUNPROTECTEDFILE & vbNewLine, "<FILE/>", Me.lblFile.Caption & vbNewLine))
140       End If
              
          'Busco la contraseña de Cada Hoja del Libro
150       If Not blnCantOpen Then
160           Call BruteForceAttackSheets
          
170           Set ws = Nothing
180           wb.Close False
190           Set wb = Nothing
200           ExcelApp.Quit
210           Set ExcelApp = Nothing
          
220           If Not IsEmpty(FinalFile) And blnRegistrado Then
230               ShellExecute Me.hwnd, vbNullString, FinalFile, vbNullString, "C:\", SW_SHOWNORMAL
240           End If
          
250           If Not blnRegistrado Then
260               MsgBox Replace(MSGWITHREGISTERED, "<FILE/>", Me.lblFile.Caption), vbInformation, AppName
270           End If
280       End If
          
290       Call Add2Status(MSGFINALIZE & vbNewLine)
300       Call Add2Status(MSGTOTALTIME & Format$(Now - dteInicio, "hh:mm:ss") & vbNewLine)
310       Call Add2Status(MSGCOMBINATIONS & lngPwdCount & vbNewLine)
320       Me.lblTimer.Caption = Empty

330      On Error GoTo 0
340      Exit Function

GetAllPasswords_Error:
350       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure GetAllPasswords of Formulario frmMain in Line " & Erl& & vbNewLine)
End Function
Private Sub Form_Load()
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Load
'---------------------------------------------------------------------------------------
Dim strIndex  As String
          
10       On Error GoTo Form_Load_Error
   
          'JAOO 20090703: Para contar las contraseñas validas encontradas
20        ValidPwdsCount = 0
   
30        Call LoadLngFiles(App.Path & "\")
40        strIndex = getCfgLngIndex
50        If strIndex <> "" Then
60            Call LoadFrmLng(Me, Val(strIndex))
70        Else
80            Call LoadFrmLng(Me, 1)
90        End If
100       Call LoadRecentFiles
          
110       Me.lblTimer.Caption = MSGADVISORY
120       Me.tmrStatus.Enabled = False
          
          'Chequea si esta Registrado
130       Call CheckRegister
140       If blnRegistrado Then
150           lblRegistrado.Visible = True
'160           Me.ToolBar.Buttons(4).Visible = False
170       Else
180           lblRegistrado.Visible = False
'190           Me.ToolBar.Buttons(4).Visible = True
200       End If

210      On Error GoTo 0
220      Exit Sub

Form_Load_Error:
230       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
''---------------------------------------------------------------------------------------
'' Procedure : Form_MouseMove
'' Author    : Administrador - JAOO
'' Date      : 20/01/2008
'' Purpose   : Manejador de Error
''---------------------------------------------------------------------------------------
'
'10       On Error GoTo Form_MouseMove_Error
'
'20       lblLnkWeb.FontUnderline = False
'30       lblMail.FontUnderline = False
'40       Screen.MousePointer = vbDefault
'
'50       On Error GoTo 0
'60       Exit Sub
'
'Form_MouseMove_Error:
'70        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Formulario frmMain in Line " & Erl& & vbNewLine)
'End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
'---------------------------------------------------------------------------------------
' Procedure : Form_OLEDragDrop
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

      Dim strAux As String

10       On Error GoTo Form_OLEDragDrop_Error

20        strAux = Data.Files(1)
30        If UCase(Right(strAux, 3)) = "XLS" Then
40            Me.lblFile.Caption = Data.Files(1)
50        End If

60       On Error GoTo 0
70       Exit Sub

Form_OLEDragDrop_Error:
80        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Form_OLEDragDrop of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Private Sub lblFile_Change()
'---------------------------------------------------------------------------------------
' Procedure : lblFile_Change
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------
Dim intI As Integer

10       On Error GoTo lblFile_Change_Error

20        Me.txtResult.Text = Empty
30        Me.lblFile.ToolTipText = Me.lblFile.Caption
40        If Me.lblFile.Caption <> Empty Then
50            For intI = 4 To 1 Step -1
60                If arr(intI - 1) = Me.lblFile.Caption Then
70                    arr(intI - 1) = Empty
80                    mnuReciente(intI - 1).Caption = Empty
90                    intI = intI + 1
100               End If
110               If intI <= 4 Then
120                   arr(intI) = arr(intI - 1)
130               End If
140           Next intI
150           arr(intI) = Me.lblFile.Caption
              
160           For intI = 0 To 4 Step 1
170               If arr(intI) <> Empty Then
180                   mnuReciente(intI).Caption = arr(intI)
190                   mnuReciente(intI).Visible = True
200               Else
210                   mnuReciente(intI).Caption = ""
220                   mnuReciente(intI).Visible = False
230               End If
240           Next intI
250       End If

260      On Error GoTo 0
270      Exit Sub

lblFile_Change_Error:
280       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure lblFile_Change of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub
Private Sub lblFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
'---------------------------------------------------------------------------------------
' Procedure : lblFile_OLEDragDrop
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

      Dim strAux As String

10       On Error GoTo lblFile_OLEDragDrop_Error

20        strAux = Data.Files(1)
30        If UCase(Right(strAux, 3)) = "XLS" Then
40            Me.lblFile.Caption = Data.Files(1)
50        End If

60       On Error GoTo 0
70       Exit Sub

lblFile_OLEDragDrop_Error:
80        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure lblFile_OLEDragDrop of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo Form_Unload_Error

20        Call SaveRecentFiles
30        End

40       On Error GoTo 0
50       Exit Sub

Form_Unload_Error:
60        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Public Sub lblLnkWeb_Click()
'---------------------------------------------------------------------------------------
' Procedure : lblLnkWeb_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

   On Error GoTo lblLnkWeb_Click_Error
   
    'JAOO 20091104: Visita la web del producto en el idioma correspondiente
1    If IsSpanish Then
2         GoToInternet Me, "http://www.xaviware.cl/pwdxls_es.aspx", Normal
3    Else
10        GoToInternet Me, "http://www.xaviware.cl/pwdxls.aspx", Normal
11   End If

   On Error GoTo 0
   Exit Sub

lblLnkWeb_Click_Error:
    WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure lblLnkWeb_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

'Private Sub lblMail_Click()
''---------------------------------------------------------------------------------------
'' Procedure : lblMail_Click
'' Author    : Administrador - JAOO
'' Date      : 20/01/2008
'' Purpose   : Manejador de Error
''---------------------------------------------------------------------------------------
'
'10       On Error GoTo lblMail_Click_Error
'
'20        GoToInternet Me, "mailto:xaviware@xaviware.cl", Normal
'
'30       On Error GoTo 0
'40       Exit Sub
'
'lblMail_Click_Error:
'50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure lblMail_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
'End Sub

Private Sub mnuAcercaDe_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuAcercaDe_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo mnuAcercaDe_Click_Error

20        frmAbout.Show vbModal

30       On Error GoTo 0
40       Exit Sub

mnuAcercaDe_Click_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAcercaDe_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Private Sub mnuAyudaF1_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuAyudaF1_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo mnuAyudaF1_Click_Error

20        ShellExecute Me.hwnd, "Open", App.Path & "\help.htm", "", "", Normal

30       On Error GoTo 0
40       Exit Sub

mnuAyudaF1_Click_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAyudaF1_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Private Sub mnuCAPE_Click()
    Call GOTOWeb("http://www.xaviware.com/cape_es.aspx", "http://www.xaviware.com/cape.aspx")
End Sub

Public Sub mnuCompra_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuCompra_Click
' Author    : Javier Orellana - XaviWare Software Ltda.
' Date      : 20/01/2008
' Purpose   : 'Redireccionar al formulario de compra
'---------------------------------------------------------------------------------------
   On Error GoTo mnuCompra_Click_Error
   
   'JAOO 20091104 : redirecciona según el idioma
   If IsSpanish Then
        GoToInternet Me, "http://www.xaviware.cl/purchase_es.aspx?productid=pwdxls", Normal
   Else
        GoToInternet Me, "http://www.xaviware.cl/purchase.aspx?productid=pwdxls", Normal
   End If

   On Error GoTo 0
   Exit Sub

mnuCompra_Click_Error:
    WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCompra_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub
Private Sub mnuFeedBack_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuFeedBack_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------
Dim strAux As String
          
10       On Error GoTo mnuFeedBack_Click_Error

20        strAux = Empty
30        strAux = strAux & "mailto:support@xaviware.cl"
40        strAux = strAux & "?SUBJECT=Feedback - " & AppName
50        GoToInternet Me, strAux, Normal

60       On Error GoTo 0
70       Exit Sub

mnuFeedBack_Click_Error:
80        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFeedBack_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub
Private Sub mnuGoToWeb_Click()
    Call lblLnkWeb_Click
End Sub

Private Sub mnuLang_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Procedure : mnuLang_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------
10       On Error GoTo mnuLang_Click_Error

20        Call LoadFrmLng(Me, Index)
30        Call SetCfgLngIndex(Index)

40       On Error GoTo 0
50       Exit Sub

mnuLang_Click_Error:
60        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLang_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Private Sub mnuPIDGIN_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdpidgin_es.aspx", "http://www.xaviware.com/pwdpidgin.aspx")
End Sub

Private Sub mnuPWDCHR_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdchr_es.aspx", "http://www.xaviware.com/pwdchr.aspx")
End Sub

Private Sub mnuPWDChromePlus_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdchromeplus_es.aspx", "http://www.xaviware.com/pwdchromeplus.aspx")
End Sub

Private Sub mnuPWDFB_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdfb_es.aspx", "http://www.xaviware.com/pwdfb.aspx")
End Sub

Private Sub mnuPWDFlock_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdflock_es.aspx", "http://www.xaviware.com/pwdflock.aspx")
End Sub

Private Sub mnuPWDFZ_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdfz_es.aspx", "http://www.xaviware.com/pwdfz.aspx")
End Sub

Private Sub mnuPWDGMAIL_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdgmail_es.aspx", "http://www.xaviware.com/pwdgmail.aspx")
End Sub

Private Sub mnuPWDGTALK_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdtrillian_es.aspx", "http://www.xaviware.com/pwdtrillian.aspx")
End Sub

Private Sub mnuPWDIE_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdie_es.aspx", "http://www.xaviware.com/pwdie.aspx")
End Sub

Private Sub mnuPWDIRON_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdiron_es.aspx", "http://www.xaviware.com/pwdiron.aspx")
End Sub

Private Sub mnuPWDMDB_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdmdb_es.aspx", "http://www.xaviware.com/pwdmdb.aspx")
End Sub

Private Sub mnuPWDMFF_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdmff_es.aspx", "http://www.xaviware.com/pwdmff.aspx")
End Sub

Private Sub mnuPWDRAGON_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdragon_es.aspx", "http://www.xaviware.com/pwdragon.aspx")
End Sub

Private Sub mnuPWDSafari_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdsafari_es.aspx", "http://www.xaviware.com/pwdsafari.aspx")
End Sub

Private Sub mnuPWDTW_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdtw_es.aspx", "http://www.xaviware.com/pwdtw.aspx")
End Sub

Private Sub mnuRockPWD_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdrockmelt_es.aspx", "http://www.xaviware.com/pwdrockmelt.aspx")
End Sub

Private Sub mnuRRPDF_Click()
    Call GOTOWeb("http://www.xaviware.com/rrpdf_es.aspx", "http://www.xaviware.com/rrpdf.aspx")
End Sub

Private Sub mnuPWDMSN_Click()
   Call GOTOWeb("http://www.xaviware.com/pwdmsn_es.aspx", "http://www.xaviware.com/pwdmsn.aspx")
End Sub

Private Sub mnuPWDTRILLIAN_Click()
    Call GOTOWeb("http://www.xaviware.com/pwdmsn_es.aspx", "http://www.xaviware.com/pwdmsn.aspx")
End Sub

Private Sub mnuReciente_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Procedure : mnuReciente_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo mnuReciente_Click_Error

20        Me.lblFile.Caption = mnuReciente(Index).Caption

30       On Error GoTo 0
40       Exit Sub

mnuReciente_Click_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure mnuReciente_Click of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Private Sub mnuRegistro_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuRegistro_Click
' Author    : Javier Orellana - XaviWare Software Ltda.
' Date      : 04/11/2009
' Purpose   : Llama al formulario de registro
'---------------------------------------------------------------------------------------
'
    frmRegistro.Show vbModal
    Call Me.VisibleByRegister 'debe recargar en caso de quedar registrado
End Sub

Private Sub tmrStatus_Timer()
      'Detección de Debuggers
      Dim hwnd As Long
      Dim i As Integer
          
10        On Error Resume Next    'Revisa con Kernel de Win
20        If IsDebuggerPresent <> 0 Then
30            End
40        End If
          
          'Llena el Arreglo con las ClassNames
50        Call FillClassName

          'Monitorea Todas las Aplicaciones que Estan Corriendo en el
          'Entorno de Windows y Cierra Aquellas cuyos ClassNames estan
          'almacenadas en el Arreglo AppClassName
60        For i = 0 To NoOfAppClassMonitored
70            hwnd = AppPresent(Trim(AppClassName(i)), frmMain)
80            If hwnd <> 0 Then
90                KillWin hwnd
100           End If
110       Next i
          
120       Me.lblTimer.Caption = MSGTOTALTIME & Format$(Now - dteInicio, "hh:mm:ss") & vbCrLf & MSGCOMBINATIONS & lngPwdCount
End Sub

'Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
''---------------------------------------------------------------------------------------
'' Procedure : ToolBar_ButtonClick
'' Author    : Administrador - JAOO
'' Date      : 20/01/2008
'' Purpose   : Manejador de Error
''---------------------------------------------------------------------------------------
'
'10       On Error GoTo ToolBar_ButtonClick_Error
'
'20        Select Case Button.Key
'              Case "Examinar"
'30                Call btnExaminar
'40            Case "Obtener"
'50                Call btnObtener
'60            Case "Salir"
'70                Call btnSalir
'80            Case "Paypal"
'90                Call btnPaypal
'100       End Select
'
'110      On Error GoTo 0
'120      Exit Sub
'
'ToolBar_ButtonClick_Error:
'130       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure ToolBar_ButtonClick of Formulario frmMain in Line " & Erl& & vbNewLine)
'End Sub

Sub btnExaminar()
10        On Error Resume Next
20        OpenFile.FileName = ""
30        OpenFile.Filter = "Excel Files (*.Xls)|*.Xls"
40        OpenFile.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
50        OpenFile.ShowOpen
60        Me.lblFile.Caption = OpenFile.FileName
70        Me.txtResult.Text = Empty
End Sub
Sub btnObtener()
'---------------------------------------------------------------------------------------
' Procedure : btnObtener
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Obtiene Contraseñas
'---------------------------------------------------------------------------------------

10       On Error GoTo btnObtener_Error

20        If Len(Me.lblFile.Caption) = 0 Then
30            MsgBox MSGSELECTFILE, vbCritical, AppName
40            Exit Sub
50        End If
          
60        If Not gFileExists(Me.lblFile.Caption) Then
70            MsgBox MSGFILENOTFOUND, vbCritical, AppName
80            Exit Sub
90        End If
              
100       Set ExcelApp = CreateObject("Excel.Application")
110       strWorkBookPwd = Empty
          
120       Me.txtResult.Text = Empty
130       Screen.MousePointer = vbHourglass
140       Me.Habilita (False)
          
150       FinalFile = Empty
160       blnCantOpen = False
170       Call GetAllPasswords
          
180       Me.Habilita (True)
190       Screen.MousePointer = vbDefault
200       Me.tmrStatus.Enabled = False

210      On Error GoTo 0
220      Exit Sub

btnObtener_Error:
230       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure btnObtener of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub
'Sub btnSalir()
''---------------------------------------------------------------------------------------
'' Procedure : btnSalir
'' Author    : Administrador - JAOO
'' Date      : 20/01/2008
'' Purpose   : Salir
''---------------------------------------------------------------------------------------
'
'10       On Error GoTo btnSalir_Error
'
'20        Unload Me
'
'30       On Error GoTo 0
'40       Exit Sub
'
'btnSalir_Error:
'50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure btnSalir of Formulario frmMain in Line " & Erl& & vbNewLine)
'End Sub

Sub Habilita(blnEnabled As Boolean)
'---------------------------------------------------------------------------------------
' Procedure : Habilita
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo Habilita_Error

20        Me.mnuOpciones.Enabled = blnEnabled
30        Me.mnuAyuda.Enabled = blnEnabled
'40        Me.ToolBar.Enabled = blnEnabled

50       On Error GoTo 0
60       Exit Sub

Habilita_Error:
70        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Habilita of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

Public Sub VisibleByRegister()
'---------------------------------------------------------------------------------------
' Procedure : VisibleByRegister
' Author    : Javier Orellana - XaviWare Software Ltda.
' Date      : 04/11/2009
' Purpose   : Habilita o Desabilita opciones según si el software esta registrado o no
'---------------------------------------------------------------------------------------
'
120       If blnRegistrado Then
130           lblRegistrado.Visible = True
131           Me.cmdBuyNow.Visible = False
    'JAOO 20091104
132           Me.mnuRegistro.Visible = False
133           Me.mnuCompra.Visible = False
134           Me.mnulin2.Visible = False
140       Else
150           lblRegistrado.Visible = False
151           Me.cmdBuyNow.Visible = True
    'JAOO 20091104
152           Me.mnuRegistro.Visible = True
153           Me.mnuCompra.Visible = True
154           Me.mnulin2.Visible = True
160       End If
End Sub

Public Sub GOTOWeb(url_spanish As String, url_english As String)
'---------------------------------------------------------------------------------------
' Procedure : GOTOWeb
' Author    : XaviWare
' Date      : 15-05-2010
' Purpose   :
'---------------------------------------------------------------------------------------
'
On Error GoTo GOTOWEB_Error
   
1    If IsSpanish Then
2         GoToInternet Me, url_spanish, Normal
3    Else
10        GoToInternet Me, url_english, Normal
11   End If

   On Error GoTo 0
   Exit Sub

GOTOWEB_Error:
    WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure GOTOWEB of Formulario frmMain in Line " & Erl& & vbNewLine)
End Sub

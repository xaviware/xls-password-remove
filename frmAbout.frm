VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   5220
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6120
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3602.938
   ScaleMode       =   0  'User
   ScaleWidth      =   5746.997
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Image ImgCabeceraSpanish 
      Height          =   1080
      Left            =   -120
      Picture         =   "frmAbout.frx":08CA
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image ImgCabeceraEnglish 
      Height          =   1080
      Left            =   -120
      Picture         =   "frmAbout.frx":1692C
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblColaborador 
      Caption         =   "Colaboradores :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   10
      Top             =   2880
      Width           =   4605
   End
   Begin VB.Label lblTranslator 
      Alignment       =   1  'Right Justify
      Caption         =   "Pamela Ubeda - "
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   240
      MouseIcon       =   "frmAbout.frx":2C990
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3120
      Width           =   3525
   End
   Begin VB.Label lblTrans 
      Caption         =   "Traductora de Interfaz"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5634.31
      Y1              =   2401.959
      Y2              =   2401.959
   End
   Begin VB.Label lblDev 
      Caption         =   "Desarrollador de Software"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Picture         =   "frmAbout.frx":2CC9A
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Javier Orellana (XaviWare) - "
      Height          =   225
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label lblDesarrolladoPor 
      Caption         =   "Desarrollado por :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   4605
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":2D2A2
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   5325
   End
   Begin VB.Label lblTitle 
      Caption         =   "Título de la aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   5325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5634.31
      Y1              =   2401.959
      Y2              =   2401.959
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión"
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   5325
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : frmAbout
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Acerca de ...
'---------------------------------------------------------------------------------------

Private Sub cmdOK_Click()
'---------------------------------------------------------------------------------------
' Procedure : cmdOK_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

10       On Error GoTo cmdOK_Click_Error

20      Unload Me

30       On Error GoTo 0
40       Exit Sub

cmdOK_Click_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOK_Click of Formulario frmAbout in Line " & Erl& & vbNewLine)
End Sub
Private Sub Form_Load()
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Manejador de Error
'---------------------------------------------------------------------------------------

Dim strIndex As String
Dim strTraductor As String
          
   On Error GoTo Form_Load_Error
   
    'JAOO 20091104 : cabecera según el idioma correspondiente
1     If IsSpanish Then
2         Me.ImgCabeceraSpanish.Visible = True
3         Me.ImgCabeceraEnglish.Visible = False
4     Else
5         Me.ImgCabeceraSpanish.Visible = False
6         Me.ImgCabeceraEnglish.Visible = True
7     End If

10        strIndex = getCfgLngIndex
20        If strIndex <> "" Then
30            Call LoadFrmLng(Me, Val(strIndex))
40        End If

50        Me.Caption = AppName
60        lblVersion.Caption = "v." & App.Major & "." & App.Minor & "." & App.Revision & " Shareware"
70        lblTitle.Caption = AppName
          
          'Cargo al Traductor
80        strTraductor = ReadINI(modLang.strLngFile, "GralInfo", "TranslatorName")
90        Me.lblTranslator = strTraductor & " -"

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
    WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmAbout in Line " & Erl& & vbNewLine)
End Sub

Private Sub lblTranslator_Click()
'---------------------------------------------------------------------------------------
' Procedure : lblTranslator_Click
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Javier Orellana - 02/10/2006 21:55 hrs.
'---------------------------------------------------------------------------------------
Dim strTransWeb As String
Dim strTransEmail As String

10       On Error GoTo lblTranslator_Click_Error

20        strTransWeb = ReadINI(modLang.strLngFile, "GralInfo", "TranslatorURL")
30        strTransEmail = ReadINI(modLang.strLngFile, "GralInfo", "TranslatorEmail")
          
40        If Trim(strTransWeb) <> Empty Then
50            GoToInternet Me, strTransWeb, Normal
60        Else
70            If Trim(strTransEmail) <> Empty Then
80                GoToInternet Me, "mailto:" & strTransEmail, Normal
90            End If
100       End If

110      On Error GoTo 0
120      Exit Sub

lblTranslator_Click_Error:
130       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure lblTranslator_Click of Formulario frmAbout in Line " & Erl& & vbNewLine)
End Sub

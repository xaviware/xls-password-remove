Attribute VB_Name = "modMain"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Modulo de Utilidades Globales
'---------------------------------------------------------------------------------------

'Variable que indica si tiene Registro Shareware
Global blnRegistrado As Boolean
Global blnCantOpen As Boolean

'Constantes
'Public Const CstrPaypalDonate = "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=sales%40xaviware%2ecl&item_name=WorkSheet%20Password%20Baypasser%20%28Donation%29&item_number=003&no_shipping=0&return=http%3a%2f%2fwww%2exaviware%2ecl&cancel_return=http%3a%2f%2fwww%2exaviware%2ecl&cn=Optional%20Comment&tax=0&currency_code=USD&bn=PP%2dDonationsBF&charset=UTF%2d8"
Public Const AppName = "Password Eliminator .XLS"
Public Const OFS_MAXPATHNAME = 128
Public Const OF_READ = &H0
Public Const SW_SHOWNORMAL = 1
Public Const MAX_LANGUAGE = 9 'Maximo de Menus Disponibles para Lenguaje
Public Const NoOfAppClassMonitored As Integer = 3
Public Const WM_CLOSE = &H10
Public Const GW_HWNDNEXT = 2
Public Const GW_OWNER = 4
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
''
Public Const LEN_PWD = &H14
Public Const LEN_PWD_CHAR = &H2
Public Const LEN_PWD_COMPILED = LEN_PWD * LEN_PWD_CHAR
Public Const LEN_DATETIME_COMPILED = &H8
''
Public Const OFFSET_PWD_INTERVAL = &H4
Public Const OFFSET_PWD_COMPILED = &H42
Public Const OFFSET_DATETIME_COMPILED = &H72
''
Public Const OFFSET_XOR_GET = OFFSET_PWD_COMPILED + LEN_PWD_COMPILED - OFFSET_PWD_INTERVAL
Public Const VAL_XOR_FIX = &H2A7DA8A8

'Variables
Public NullDate As Date
Public AppClassName(NoOfAppClassMonitored) As String
Public FinalFile As String
''Variables de Mesajes de Error y Otros
Public MSGADVISORY As String
Public MSGSELECTFILE As String
Public MSGFILENOTFOUND As String
Public MSGWITHREGISTERED As String

Public MSGREGISTERED  As String
Public MSGERRINVALIDKEY As String

Public MSGUNPROTECTEDFILE As String
Public MSGUNPROTECTEDSHEET As String

Public MSGFINALIZE As String
Public MSGINITIALIZE As String

Public MSGTOTALTIME As String
Public MSGCOMBINATIONS As String

Public MSGUNPROTECBOOK As String
Public MSGUNPROTECSHEET As String

Public MSGBOOKOK As String
Public MSGSHEETOK As String

Public MSGBOOKWARRANTY As String
Public MSGSHEETWARRANTY As String

Public MSGBOOKPWDNOTFOUND As String
Public MSGSHEETPWDNOTFOUND As String

Public MSGCANTHELPU As String
Public MSGONLYREMOVE As String
''Variable Arreglo de Archivos Recientes
Public arr(4) As String

'Apis
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hfile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
''
Public Declare Function IsDebuggerPresent Lib "kernel32" () As Long
''
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

'Tipos
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Enum T_WindowStyle
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ReadINI(File As String, AppName As String, KeyName As String) As String
'---------------------------------------------------------------------------------------
' Procedure : ReadINI
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Lee .INI
'---------------------------------------------------------------------------------------
Dim value As String * 1024
Dim i As Variant
        
10       On Error GoTo ReadINI_Error

20        i = GetPrivateProfileString(AppName, KeyName, "", value, 512, File)
30        ReadINI = Left(value, InStr(value, Chr(0)) - 1)

40       On Error GoTo 0
50       Exit Function

ReadINI_Error:
60        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure ReadINI of Módulo modMain in Line " & Erl& & vbNewLine)
End Function

Function WriteINI(File As String, AppName As String, KeyName As String, NewValue As String) As Long
'---------------------------------------------------------------------------------------
' Procedure : WriteINI
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Escribe .INI
'---------------------------------------------------------------------------------------

10       On Error GoTo WriteINI_Error

20        WriteINI = WritePrivateProfileString(AppName, KeyName, NewValue, File)

30       On Error GoTo 0
40       Exit Function

WriteINI_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure WriteINI of Módulo modMain in Line " & Erl& & vbNewLine)
End Function

Public Function gLetraCapital(ByVal pTexto As String, Optional ByVal Solo_Primero As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : gLetraCapital
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : ' Por Cristian Cabello Diaz (ZyKoN) Octubre 2002
'             ' http://www.xaviware.es.vg
'             ' Esta función cambia un string a minúsculas pero con las primeras letras de cada
'             ' palabra en mayúsculas.
'             ' Si el parámetro Solo_Primero es verdadero entonces se cambia a mayúsculas sólo
'             ' la primera letra del string
'             ' se declaran variables
'---------------------------------------------------------------------------------------
Dim blnCambiar As Boolean
Dim x As Long

10       On Error GoTo gLetraCapital_Error

          ' se asigna valor inicial
20        blnCambiar = True
          ' se convierte el texto a minúsculas.
30        pTexto = CStr(LCase(pTexto))
          ' ahora por cada letra del string se busca los que se convertirán a mayúsculas
40        For x = 1 To Len(pTexto)
              ' Si el caracter posecionado es un espacio o un punto asignamos el valor
              ' verdadero a la variable blnCambiar para señalar que el siguiente caracter
              ' se debe convertir a mayúsculas. Siempre y cuando el parámetro Solo_Primero
              ' tenga el valor falso.
50            If (Mid(pTexto, x, 1) = " " Or Mid(pTexto, x, 1) = ".") And (Not Solo_Primero) Then
60                blnCambiar = True
70            End If
              ' si la variable blnCambiar es verdadera y el caracter posecionado no es
              ' espacio ni punto podemos cambiar a mayúsculas y asignamos un falso a
              ' la variable blnCambiar, para que no cambie el siguiente caracter si es
              ' una letra.
80            If blnCambiar And Mid(pTexto, x, 1) <> " " And Mid(pTexto, x, 1) <> "." Then
90                Mid(pTexto, x, 1) = UCase(Mid(pTexto, x, 1))
100               blnCambiar = False
110           End If
120       Next
          ' asignamos resultado a la función
130       gLetraCapital = pTexto

140      On Error GoTo 0
150      Exit Function

gLetraCapital_Error:
160       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure gLetraCapital of Módulo modMain in Line " & Erl& & vbNewLine)
End Function

Function gEncripta(strCadenaEnc As String, strValor As String) As String
'---------------------------------------------------------------------------------------
' Procedure : gEncripta
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'recibe una variable tipo texto y la Encripta
'             'retorna la cadena ya encriptada
'             'strCadenaEnc = Cadena de Encriptación
'             'strValor = Teto a Encriptar
'---------------------------------------------------------------------------------------
Dim strSalida As String
Dim intI As Integer
Dim intAscii As Integer

10       On Error GoTo gEncripta_Error

20        strSalida = ""
30        For intI = 1 To Len(strValor)
40            intAscii = ((Asc(Mid(strCadenaEnc, intI, 1)) + Asc(Mid(strValor, intI, 1))) Mod 95) + 32
50            strSalida = strSalida & Chr(intAscii)
60        Next intI
70        gEncripta = strSalida

80       On Error GoTo 0
90       Exit Function

gEncripta_Error:
100       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure gEncripta of Módulo modMain in Line " & Erl& & vbNewLine)
End Function

Sub LoadRecentFiles()
      '---------------------------------------------------------------------------------------
      ' Procedure : LoadRecentFiles
      ' Author    : Administrador - JAOO
      ' Date      : 20/01/2008
      ' Purpose   : 'Javier Orellana 16-08-2004 17:33 hrs.
      '             'Carga los Archivos Recientemente abiertos por la aplicación
      '---------------------------------------------------------------------------------------
      Dim intI As Integer
      Dim strAux As String

10        On Error Resume Next
20        For intI = 0 To 4 Step 1
30            strAux = ReadINI(App.Path & "\config.cfg", "Reciente", "Rec" & intI)
40            If strAux <> "" Then
50                frmMain.mnuReciente(intI).Caption = strAux
60                arr(intI) = strAux
70            Else
80                frmMain.mnuReciente(intI).Visible = False
90            End If
100       Next intI
End Sub

Sub SaveRecentFiles()
'---------------------------------------------------------------------------------------
' Procedure : SaveRecentFiles
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Javier Orellana 17-08-2004 09:24 hrs.
'             'Guarda los Archivos Recientes en la Configuración
'---------------------------------------------------------------------------------------
Dim intI As Integer

10       On Error GoTo SaveRecentFiles_Error

20        For intI = 0 To 4 Step 1
30            Call WriteINI(App.Path & "\config.cfg", "Reciente", "Rec" & intI, arr(intI))
40        Next intI

50       On Error GoTo 0
60       Exit Sub

SaveRecentFiles_Error:
70        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure SaveRecentFiles of Módulo modMain in Line " & Erl& & vbNewLine)
End Sub
'' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GoToInternet(Parent As Form, URL As String, WindowStyle As T_WindowStyle)
      '---------------------------------------------------------------------------------------
      ' Procedure : GoToInternet
      ' Author    : Administrador - JAOO
      ' Date      : 20/01/2008
      ' Purpose   : 'Javier Orellana 19-09-2005 20:30 hrs.
      '             'Abre el Browser por Defecto con la Pagina URL
      '---------------------------------------------------------------------------------------
      Dim lngReturn As Long
          
10        On Error Resume Next
20        If fsGetBrowserInfo Then
30            lngReturn = ShellExecute(Parent.hwnd, "Open", URL, "", "", WindowStyle)
40            If lngReturn <= 32 Then ' Error
                  'MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
50            End If
60        Else
              '* nope no Browser
              'MsgBox "Could not find associated Browser", vbExclamation, "Browser Not Found"
70        End If
End Sub

Public Function fsGetBrowserInfo() As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : fsGetBrowserInfo
      ' Author    : Administrador - JAOO
      ' Date      : 20/01/2008
      ' Purpose   : 'Javier Orellana 19-09-2005 20:30 hrs.
      '             'Detecta el Browser por Defecto
      '---------------------------------------------------------------------------------------
      Dim strFileName As String, strDummy As String
      Dim strBrowserExec As String * 255
      Dim lngRetVal As Long
      Dim intFileNumber As Integer
          
10        On Error Resume Next

          ' Create a HTML file
20        strBrowserExec = Space(255)
30        strFileName = App.Path + "\TEST.HTM"
40        intFileNumber = FreeFile
50        Open strFileName For Output As #intFileNumber
60        Write #intFileNumber, "<HTML> <\HTML>"
70        Close #intFileNumber
          ' Do we have an associated application?
80        lngRetVal = FindExecutable(strFileName, strDummy, strBrowserExec)
90        strBrowserExec = Trim(strBrowserExec)
          ' If an application is found return True
100       If lngRetVal <= 32 Or strBrowserExec = "" Then ' Error
110           fsGetBrowserInfo = False
120       Else
130           fsGetBrowserInfo = True
140       End If
150       Kill strFileName
End Function

Public Function gFileExists(sFullPath As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : gFileExists
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Verifica si existe un archivo
'             'Microsoft Scripting Runtime en las Referencias de Visual Basic.
'             '2007-04-24
'---------------------------------------------------------------------------------------
Dim oFile As New Scripting.FileSystemObject
          
10       On Error GoTo gFileExists_Error

20        gFileExists = oFile.FileExists(sFullPath)

30       On Error GoTo 0
40       Exit Function

gFileExists_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure gFileExists of Módulo modMain in Line " & Erl& & vbNewLine)
End Function

Public Sub WriteLogToFile(strData As String)
      '---------------------------------------------------------------------------------------
      ' Procedure : WriteLogToFile
      ' Author    : JAOO
      ' Date      : 03/12/2007
      ' Purpose   : Registra en Log de Errores
      '---------------------------------------------------------------------------------------
      Dim intFreeFile As Integer
          
10        MsgBox strData & "please, report this problem to support@xaviware.cl", vbCritical, AppName
          
20        intFreeFile = FreeFile
30        Open App.Path & "\err.log" For Append As intFreeFile
40        Print #intFreeFile, strData
50        Close intFreeFile
End Sub

Public Function IsSpanish() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsSpanish
' Author    : Javier Orellana - XaviWare Software Ltda.
' Date      : 30/05/2010
' Purpose   : Saber si la pantalla esta en español en este momento
'---------------------------------------------------------------------------------------
'
    IsSpanish = (frmMain.mnuOpciones.Caption = "&Opciones")
End Function

Sub Main()
10        frmMain.Show
End Sub

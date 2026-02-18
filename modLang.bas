Attribute VB_Name = "modLang"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : modLang
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Modulo para Manejo de Idioma de Interfaz
'---------------------------------------------------------------------------------------
Global strLngFile As String

Sub LoadFrmLng(frm As Form, Index As Integer)
      '---------------------------------------------------------------------------------------
      ' Procedure : LoadFrmLng
      ' Author    : Administrador - JAOO
      ' Date      : 20/01/2008
      ' Purpose   : 'Javier Orellana 14/08/2004 17:00 Hrs.
      '             'Carga el Idioma al Formulario
      '---------------------------------------------------------------------------------------
      Dim ctlCtrl As Control
      Dim strTipo As String
      Dim strAux As String
      Dim intI As Integer

10        On Error Resume Next
20        strLngFile = App.Path & "\" & frmMain.mnuLang(Index).Caption & ".lng"
30        frmMain.mnuLang(Index).Checked = True
40        For intI = 0 To MAX_LANGUAGE Step 1
50            If intI <> Index Then
60                frmMain.mnuLang(intI).Checked = False
70            End If
80        Next intI
       
90        For Each ctlCtrl In frm.Controls
100           strTipo = TypeName(ctlCtrl)
110           If strTipo = "Label" Or strTipo = "Menu" Or strTipo = "CommandButton" Or strTipo = "Frame" Then
120               strAux = ReadINI(strLngFile, ctlCtrl.Name, "Caption")
130               If Len(Trim(strAux)) <> 0 Then
140                   ctlCtrl.Caption = Replace(strAux, "<APP/>", AppName)
150               End If

                'JAOO 20100614 - Incluir ToolTipText de Botones
151             strAux = ReadINI(strLngFile, ctlCtrl.Name, "ToolTipText")
152             If Len(Trim(strAux)) <> 0 Then
153                    ctlCtrl.ToolTipText = Replace(strAux, "<APP/>", AppName)
154             End If

160           End If
170       Next
'180       Call LoadToolBarLng(frmMain.ToolBar, Index)
190       Call LoadMsgLng(Index)
End Sub
'Sub LoadToolBarLng(ToolBar As ToolBar, Index As Integer)
''---------------------------------------------------------------------------------------
'' Procedure : LoadToolBarLng
'' Author    : Administrador - JAOO
'' Date      : 20/01/2008
'' Purpose   : 'Javier Orellana 22/08/2004 18:32 Hrs.
''             'Carga el Idioma al Formulario
''---------------------------------------------------------------------------------------
'Dim btn As Button
'Dim strAux As String
'Dim strLngFile As String
'
'10       On Error GoTo LoadToolBarLng_Error
'
'20        strLngFile = App.Path & "\" & frmMain.mnuLang(Index).Caption & ".lng"
'30        For Each btn In frmMain.ToolBar.Buttons
'40            strAux = ReadINI(strLngFile, btn.Key, "ToolTipText")
'50            If Len(Trim(strAux)) <> 0 Then
'60                btn.ToolTipText = strAux
'70            End If
'80        Next
'
'90       On Error GoTo 0
'100      Exit Sub
'
'LoadToolBarLng_Error:
'110       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure LoadToolBarLng of Módulo modLang in Line " & Erl& & vbNewLine)
'End Sub
Sub LoadMsgLng(Index As Integer)
'---------------------------------------------------------------------------------------
' Procedure : LoadMsgLng
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Javier Orellana 16/08/2004 11:06 Hrs.
'             'Carga los mensajes de la aplicaciónn
'---------------------------------------------------------------------------------------
Dim strAux As String
Dim strLngFile As String

10       On Error GoTo LoadMsgLng_Error

20        strLngFile = App.Path & "\" & frmMain.mnuLang(Index).Caption & ".lng"
          
          'Cargo los por defecto
      '    MSGADVISORY = "Este programa, solo recupera contraseñas de las hojas de un libro de excel no protegido."
      '    MSGSELECTFILE = "No ha Seleccionado un Archivo..."
      '    MSGFILENOTFOUND = "Archivo no Encontrado..."
      '    MSGINITIALIZE = "Proceso de Desprotección Inicializado..."
      '    MSGFINALIZE = "Proceso de Desprotección Finalizado..."
      '    MSGUNPROTECTEDFILE = "El Archivo <FILE/> No esta protegido con Contraseña"
      '    MSGUNPROTECTEDSHEET = "La Hoja <SHEET/> No esta protegida con Contraseña"
      '    MSGREGISTERED = "<APP/> ha sido registrado correctamente"
      '    MSGTOTALTIME = "Tiempo Total : "
      '    MSGCOMBINATIONS = "Total de Combinaciones Probadas : "
      '    MSGUNPROTECBOOK = "Desprotegiendo el Libro <BOOK/>"
      '    MSGUNPROTECSHEET = "Desprotegiendo la Hoja <SHEET/>"
      '    MSGBOOKOK = "el Libro <BOOK/> Ha sido Desprotegido Satisfactoriamente"
      '    MSGSHEETOK = "La Hoja <SHEET/> ha sido Desprotegida Satisfactoriamente"
      '    MSGBOOKWARRANTY = "Este mensaje es una garantia de que el Libro <BOOK/> puede ser Desprotegido"
      '    MSGSHEETWARRANTY = "Este mensaje es una garantia de que La Hoja <SHEET/> puede ser Desprotegida"
      '    MSGBOOKPWDNOTFOUND = "ADVERTENCIA: Una Combinación válida para el Libro <BOOK/> No Ha Sido Encontrada."
      '    MSGSHEETPWDNOTFOUND = "ADVERTENCIA: Una Combinación válida para la Hoja <SHEET/> No Ha Sido Encontrada."
      '    MSGERRInvalidKey = "La clave de registro especificada no es válida, por favor intente nuevamente"
      '    MSGWITHREGISTERED = "Con la compra de <APP/> podrá abrir una versión desprotegida del archivo <FILE/>"
      '    MSGCANTHELPU = "Este programa no puede quitar la protección de archivo, Este programa no puede ayudarte con este Archivo."
      '    MSGONLYREMOVE = "Este programa solo remueve contraseñas establecidas en [Herramientas]-[Proteger]-[Proteger Hoja]/[Proteger Libro]."
          
30        strAux = ReadINI(strLngFile, "MSG", "MSGADVISORY")
40        If Len(Trim(strAux)) <> 0 Then MSGADVISORY = strAux
          
50        strAux = ReadINI(strLngFile, "MSG", "MSGSELECTFILE")
60        If Len(Trim(strAux)) <> 0 Then MSGSELECTFILE = strAux
          
70        strAux = ReadINI(strLngFile, "MSG", "MSGFILENOTFOUND")
80        If Len(Trim(strAux)) <> 0 Then MSGFILENOTFOUND = strAux
          
90        strAux = ReadINI(strLngFile, "MSG", "MSGINITIALIZE")
100       If Len(Trim(strAux)) <> 0 Then MSGINITIALIZE = strAux
          
110       strAux = ReadINI(strLngFile, "MSG", "MSGFINALIZE")
120       If Len(Trim(strAux)) <> 0 Then MSGFINALIZE = strAux
          
130       strAux = ReadINI(strLngFile, "MSG", "MSGUNPROTECTEDFILE")
140       If Len(Trim(strAux)) <> 0 Then MSGUNPROTECTEDFILE = strAux
          
150       strAux = ReadINI(strLngFile, "MSG", "MSGUNPROTECTEDSHEET")
160       If Len(Trim(strAux)) <> 0 Then MSGUNPROTECTEDSHEET = strAux
          
170       strAux = ReadINI(strLngFile, "MSG", "MSGREGISTERED")
180       If Len(Trim(strAux)) <> 0 Then MSGREGISTERED = Replace(strAux, "<APP/>", AppName)
          
190       strAux = ReadINI(strLngFile, "MSG", "MSGERRINVALIDKEY")
200       If Len(Trim(strAux)) <> 0 Then MSGERRINVALIDKEY = strAux
          
210       strAux = ReadINI(strLngFile, "MSG", "MSGTOTALTIME")
220       If Len(Trim(strAux)) <> 0 Then MSGTOTALTIME = strAux
          
230       strAux = ReadINI(strLngFile, "MSG", "MSGCOMBINATIONS")
240       If Len(Trim(strAux)) <> 0 Then MSGCOMBINATIONS = strAux
          
250       strAux = ReadINI(strLngFile, "MSG", "MSGUNPROTECBOOK")
260       If Len(Trim(strAux)) <> 0 Then MSGUNPROTECBOOK = strAux
          
270       strAux = ReadINI(strLngFile, "MSG", "MSGUNPROTECSHEET")
280       If Len(Trim(strAux)) <> 0 Then MSGUNPROTECSHEET = strAux
          
290       strAux = ReadINI(strLngFile, "MSG", "MSGBOOKOK")
300       If Len(Trim(strAux)) <> 0 Then MSGBOOKOK = strAux
          
310       strAux = ReadINI(strLngFile, "MSG", "MSGSHEETOK")
320       If Len(Trim(strAux)) <> 0 Then MSGSHEETOK = strAux
          
330       strAux = ReadINI(strLngFile, "MSG", "MSGBOOKWARRANTY")
340       If Len(Trim(strAux)) <> 0 Then MSGBOOKWARRANTY = strAux
          
350       strAux = ReadINI(strLngFile, "MSG", "MSGSHEETWARRANTY")
360       If Len(Trim(strAux)) <> 0 Then MSGSHEETWARRANTY = strAux
          
370       strAux = ReadINI(strLngFile, "MSG", "MSGBOOKPWDNOTFOUND")
380       If Len(Trim(strAux)) <> 0 Then MSGBOOKPWDNOTFOUND = strAux
          
390       strAux = ReadINI(strLngFile, "MSG", "MSGSHEETPWDNOTFOUND")
400       If Len(Trim(strAux)) <> 0 Then MSGSHEETPWDNOTFOUND = strAux
          
410       strAux = ReadINI(strLngFile, "MSG", "MSGWITHREGISTERED")
420       If Len(Trim(strAux)) <> 0 Then MSGWITHREGISTERED = Replace(strAux, "<APP/>", AppName)
          
430       strAux = ReadINI(strLngFile, "MSG", "MSGCANTHELPU")
440       If Len(Trim(strAux)) <> 0 Then MSGCANTHELPU = strAux
          
450       strAux = ReadINI(strLngFile, "MSG", "MSGONLYREMOVE")
460       If Len(Trim(strAux)) <> 0 Then MSGONLYREMOVE = strAux

470      On Error GoTo 0
480      Exit Sub

LoadMsgLng_Error:
490       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure LoadMsgLng of Módulo modLang in Line " & Erl& & vbNewLine)
End Sub
Sub LoadLngFiles(strPath As String)
      'XaviWare(c). 2004' Javier Orellana
      'Carga los Archivos de Idioma al Menu
      Dim strFile As String
      Dim intCont As Integer
      Dim intI As Integer

10        On Error Resume Next
20        intCont = 0
30        strFile = Dir(strPath & "*.lng", vbArchive)
40        Do While strFile <> ""
50            frmMain.mnuLang(intCont).Caption = gLetraCapital(Mid(strFile, 1, (Len(strFile) - 4)))
60            intCont = intCont + 1
70            strFile = Dir()
80        Loop
          
          'Visible False, para los no asignados
90        For intI = intCont To MAX_LANGUAGE Step 1
100           frmMain.mnuLang(intI).Visible = False
110       Next intI
End Sub
Function getCfgLngIndex() As String
'---------------------------------------------------------------------------------------
' Procedure : getCfgLngIndex
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Obtiene Indice de Idioma
'---------------------------------------------------------------------------------------

10       On Error GoTo getCfgLngIndex_Error

20        getCfgLngIndex = ReadINI(App.Path & "\config.cfg", "Idioma", "Index")

30       On Error GoTo 0
40       Exit Function

getCfgLngIndex_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure getCfgLngIndex of Módulo modLang in Line " & Erl& & vbNewLine)
End Function
Sub SetCfgLngIndex(Index As Integer)
'---------------------------------------------------------------------------------------
' Procedure : SetCfgLngIndex
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Setea Indice de Idioma
'---------------------------------------------------------------------------------------

10       On Error GoTo SetCfgLngIndex_Error

20        Call WriteINI(App.Path & "\config.cfg", "Idioma", "Index", Str(Index))

30       On Error GoTo 0
40       Exit Sub

SetCfgLngIndex_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure SetCfgLngIndex of Módulo modLang in Line " & Erl& & vbNewLine)
End Sub

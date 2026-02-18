Attribute VB_Name = "modGetPwd"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : modGetPwd
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Modulo Principal para obtención de contraseñas
'---------------------------------------------------------------------------------------
Global ExcelApp As Object          'Object de Excel
Global wb As Object                'As excel.Workbook  'Workbook de Excel
Global ws As Object                'As Worksheet       'Worksheet de Excel

Global lngPwdCount As Long         'Para Contar la Cantidad de Password Probadas
Global dteInicio As Date           'Para Indicar el Tiempo Transcurrido

Global strWorkBookPwd As String    'Almacena la Contraseña del Libro
Global strCurrentPwd As String     'Contraseña Que se esta probando Actualmente

'JAOO 20090703: Recuperar en base a las ya encontradas
Global ValidPwds() As String       'Almacena las contraseñas validas encontradas en la ejecuciòn actual
Global ValidPwdsCount As Integer   'Contador de las contraseñas valiudas encontradas en la ejecuciòn actual

Public Sub Add2ValidPwds(strPwd As String)
'---------------------------------------------------------------------------------------
' Procedure : Add2ValidPwds
' Author    : Administrador - JAOO
' Date      : 03/07/2009
' Purpose   : Agrega la Password Especificada a la lista de contraseñas validas
'---------------------------------------------------------------------------------------
'
Dim i As Integer

10      On Error GoTo Add2ValidPwds_Error

20      If ValidPwdsCount > 0 Then
30          For i = 0 To ValidPwdsCount - 1
40              If strPwd = ValidPwds(i) Then
50                  Exit Sub 'No se Agrega si ya existe
60              End If
70          Next i
80      End If

90      ValidPwdsCount = ValidPwdsCount + 1
100     ReDim Preserve ValidPwds(ValidPwdsCount)
          
110     ValidPwds(ValidPwdsCount - 1) = strPwd

120     On Error GoTo 0
130     Exit Sub

Add2ValidPwds_Error:

140    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Add2ValidPwds of Módulo modGetPwd"
End Sub

Public Sub Add2Status(strMsg As String)
'---------------------------------------------------------------------------------------
' Procedure : Add2Status
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Agrega un Mensaje al Cuadro de Resultados
'---------------------------------------------------------------------------------------

10       On Error GoTo Add2Status_Error

20        frmMain.txtResult.Text = frmMain.txtResult.Text & strMsg

30       On Error GoTo 0
40       Exit Sub

Add2Status_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Add2Status of Módulo modGetPwd in Line " & Erl& & vbNewLine)
End Sub

Public Sub BruteForceAttackSheets()
'---------------------------------------------------------------------------------------
' Procedure : BruteForceAttackSheets
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Recupera las Contraseñas de Todas las Hojas del Libro
'---------------------------------------------------------------------------------------
Dim i As Integer

10       On Error GoTo BruteForceAttackSheets_Error

          'Recorro todas las Hojas
20        For i = 1 To wb.Worksheets.Count
30            Set ws = wb.Worksheets(i)
              'Chequeo si la Hoja Tiene Contraseña
40            If CheckPwdSheetSet Then
                  'Busco la Contraseña de la Hoja
45                If Not modGetPwd.BasedOnValidPwdsAttackSheet Then
50                      BruteForceAttackSheet
55                End If
60            Else
70                Call Add2Status(Replace(MSGUNPROTECTEDSHEET & vbNewLine, "<SHEET/>", ws.Name))
80            End If
90        Next

100      On Error GoTo 0
110      Exit Sub

BruteForceAttackSheets_Error:
120       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure BruteForceAttackSheets of Módulo modGetPwd in Line " & Erl& & vbNewLine)
End Sub

Public Function TryWithSheet(strPwd As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : TryWithSheet
' Author    : Administrador - JAOO
' Date      : 03/07/2009
' Purpose   : Hace las operaciones de un intento con una contraseña para la Hoja
'---------------------------------------------------------------------------------------
'
10       On Error GoTo TryWithSheet_Error

20        TryWithSheet = False
30        If FindPwdSheet(strPwd) Then
40            If blnRegistrado Then
50                Call SaveUnprotectedCopy
60                Call Add2Status(Replace(MSGSHEETOK & vbNewLine, "<SHEET/>", ws.Name))
70            Else
80                Call Add2Status(Replace(MSGSHEETWARRANTY & vbNewLine, "<SHEET/>", ws.Name))
90            End If
100           TryWithSheet = True
110           Exit Function
120       End If
130       DoEvents

140      On Error GoTo 0
150      Exit Function

TryWithSheet_Error:

160       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TryWithSheet of Módulo modGetPwd"
End Function

Public Sub BruteForceAttackSheet()
      Dim i As Integer, j As Integer, k As Integer
      Dim l As Integer, m As Integer, n As Integer
      Dim i1 As Integer, i2 As Integer, i3 As Integer
      Dim i4 As Integer, i5 As Integer, i6 As Integer
      Dim strPwd As String
         
10        On Error Resume Next
20        Add2Status (Replace(MSGUNPROTECSHEET, "<SHEET/>", ws.Name) & vbNewLine)
30        For i = 65 To 66
40            For j = 65 To 66
50                For k = 65 To 66
60                    For l = 65 To 66
70                        For m = 65 To 66
80                            For i1 = 65 To 66
90                                For i2 = 65 To 66
100                                   For i3 = 65 To 66
110                                       For i4 = 65 To 66
120                                           For i5 = 65 To 66
130                                               For i6 = 65 To 66
140                                                   For n = 32 To 126
150                                                       strPwd = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
160                                                       If TryWithSheet(strPwd) Then
170                                                           Exit Sub
180                                                       End If
190                                                   Next
200                                               Next
210                                           Next
220                                       Next
230                                   Next
240                               Next
250                           Next
260                       Next
270                   Next
280               Next
290           Next
300       Next
310       Add2Status (Replace(MSGSHEETPWDNOTFOUND & vbNewLine, "<SHEET/>", ws.Name))
End Sub

Public Function BasedOnValidPwdsAttackSheet() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : BasedOnValidPwdsAttackSheet
' Author    : Administrador - JAOO
' Date      : 03/07/2009
' Purpose   : Prueba abrir la Hoja con la lista de contraseñas previamente encontradas
'             en la actual ejecuciòn de la aplicaciòn
'---------------------------------------------------------------------------------------
'
Dim i As Integer

   On Error GoTo BasedOnValidPwdsAttackSheet_Error

10        BasedOnValidPwdsAttackSheet = False
          
20        For i = 0 To ValidPwdsCount - 1 Step 1
25            If TryWithSheet(ValidPwds(i)) Then
30                BasedOnValidPwdsAttackSheet = True
32                Exit For
35            End If
40        Next i

   On Error GoTo 0
   Exit Function

BasedOnValidPwdsAttackSheet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BasedOnValidPwdsAttackSheet of Módulo modGetPwd"
End Function

Public Function BasedOnValidPwdsAttackWorkbook() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : BasedOnValidPwdsAttackWorkbook
' Author    : Administrador - JAOO
' Date      : 03/07/2009
' Purpose   : Prueba abrir el libro con la lista de contraseñas previamente encontradas
'             en la actual ejecuciòn de la aplicaciòn
'---------------------------------------------------------------------------------------
'
Dim i As Integer

10       On Error GoTo BasedOnValidPwdsAttackWorkbook_Error
        
20       BasedOnValidPwdsAttackWorkbook = False

30       For i = 0 To ValidPwdsCount - 1 Step 1
35            If TryWithWorkbook(ValidPwds(i)) Then
40                BasedOnValidPwdsAttackWorkbook = True
42                Exit For
45            End If
50       Next i

60       On Error GoTo 0
70       Exit Function

BasedOnValidPwdsAttackWorkbook_Error:

80       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BasedOnValidPwdsAttackWorkbook of Módulo modGetPwd"
End Function

Public Function TryWithWorkbook(strPwd As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : TryWithWorkbook
' Author    : Administrador - JAOO
' Date      : 03/07/2009
' Purpose   : Hace las operaciones de un intento con una contraseña para el Libro
'---------------------------------------------------------------------------------------
'
10       On Error GoTo TryWithWorkbook_Error

20        TryWithWorkbook = False
30        If FindPwdWorkBook(strPwd) Then
40            If blnRegistrado Then
50                Call SaveUnprotectedCopy
60                Call Add2Status(Replace(MSGBOOKOK & vbNewLine, "<BOOK/>", wb.Name))
70            Else
80                Call Add2Status(Replace(MSGBOOKWARRANTY & vbNewLine, "<BOOK/>", wb.Name))
90            End If
100           TryWithWorkbook = True
110           Exit Function
120       End If
130       DoEvents

140      On Error GoTo 0
150      Exit Function

TryWithWorkbook_Error:

160       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TryWithWorkbook of Módulo modGetPwd"
End Function

Public Sub BruteForceAttackWorkBook()
      Dim i As Integer, j As Integer, k As Integer
      Dim l As Integer, m As Integer, n As Integer
      Dim i1 As Integer, i2 As Integer, i3 As Integer
      Dim i4 As Integer, i5 As Integer, i6 As Integer
      Dim strPwd As String
         
10        On Error Resume Next
20        Add2Status (Replace(MSGUNPROTECBOOK, "<BOOK/>", wb.Name) & vbNewLine)
30        For i = 65 To 66
40            For j = 65 To 66
50                For k = 65 To 66
60                    For l = 65 To 66
70                        For m = 65 To 66
80                            For i1 = 65 To 66
90                                For i2 = 65 To 66
100                                   For i3 = 65 To 66
110                                       For i4 = 65 To 66
120                                           For i5 = 65 To 66
130                                               For i6 = 65 To 66
140                                                   For n = 32 To 126
150                                                       strPwd = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
160                                                       If TryWithWorkbook(strPwd) Then
170                                                           Exit Sub
180                                                       End If
190                                                   Next
200                                               Next
210                                           Next
220                                       Next
230                                   Next
240                               Next
250                           Next
260                       Next
270                   Next
280               Next
290           Next
300       Next
310       Add2Status (Replace(MSGSHEETPWDNOTFOUND & vbNewLine, "<BOOK/>", wb.Name))
End Sub

Public Sub SaveUnprotectedCopy()
'---------------------------------------------------------------------------------------
' Procedure : SaveUnprotectedCopy
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : '2007-04-25
'---------------------------------------------------------------------------------------
10        On Error GoTo SaveUnprotectedCopy_Error

20        FinalFile = Replace(frmMain.lblFile.Caption, wb.Name, Empty) & "Unprotected Copy of " & wb.Name
30        wb.SaveCopyAs (FinalFile)

40       On Error GoTo 0
50       Exit Sub

SaveUnprotectedCopy_Error:
60        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure SaveUnprotectedCopy of Módulo modGetPwd in Line " & Erl& & vbNewLine)
End Sub

Public Function FindPwdWorkBook(ByVal strPwd As String) As Boolean
      'Chequea una contraseña contra el WorkBook Excel
10    On Error GoTo NotFound

20        strCurrentPwd = strPwd
30        lngPwdCount = lngPwdCount + 1
          
40        wb.Unprotect strPwd
50        If wb.ProtectStructure = False And wb.ProtectWindows = False Then
60            FindPwdWorkBook = True
              'JAOO 20090703: Agrega la contraseña encontrada a la lista de PWDs Validas
70            Add2ValidPwds (strPwd)
80        End If
90        Exit Function

NotFound:
100       FindPwdWorkBook = False
End Function

Public Function FindPwdSheet(ByVal strPwd As String) As Boolean
      'Chequea una contraseña contra la Hoja Actual
10    On Error GoTo NotFound
          
20        strCurrentPwd = strPwd
30        lngPwdCount = lngPwdCount + 1
40        ws.Unprotect strPwd
50        FindPwdSheet = True
          
          'JAOO 20090703: Agrega la contraseña encontrada a la lista de PWDs Validas
60        Add2ValidPwds (strPwd)
70        Exit Function

NotFound:
80        FindPwdSheet = False
End Function

Public Function CheckPwdFileSet() As Boolean
      'Chequea si el Workbook Tiene o no Contraseña de
      'ProtectStructure & ProtectWindow
10    On Error GoTo Err
          
20        Set wb = ExcelApp.Workbooks.Open(frmMain.lblFile.Caption, , True, , Empty)
          
30        CheckPwdFileSet = wb.ProtectStructure
40        If Not wb.ProtectStructure Then
50            CheckPwdFileSet = wb.ProtectWindow
60        End If
70        Exit Function

Err:
80        If Err.Number = 1004 Then
90            Call Add2Status(MSGCANTHELPU & vbNewLine)
100           Call Add2Status(MSGONLYREMOVE & vbNewLine)
110           blnCantOpen = True
120           CheckPwdFileSet = True
130       Else
140           CheckPwdFileSet = False
150       End If
End Function

Public Function CheckPwdSheetSet() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : CheckPwdSheetSet
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Chequea Si la Hoja Actual Tiene o no Contraseña
'---------------------------------------------------------------------------------------
10       On Error GoTo CheckPwdSheetSet_Error

20        CheckPwdSheetSet = Not FindPwdSheet(Empty)

30       On Error GoTo 0
40       Exit Function

CheckPwdSheetSet_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure CheckPwdSheetSet of Módulo modGetPwd in Line " & Erl& & vbNewLine)
End Function

Attribute VB_Name = "modKeyGen"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : modKeyGen
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Clave de Registro
'---------------------------------------------------------------------------------------

Function KeyGen(kNamev As Variant, kPass As String, kType As Integer) As String
      '****************************************************************************
      '* KeyGen v2.01 Build 01                                                    *
      '* Copyright © 2000 W.G.Griffiths                                           *
      '* E-Mail: w.g.griffiths@telinco.co.uk                                      *
      '****************************************************************************
10    On Error Resume Next         'still here just as a precaution

      Dim cTable(512) As Integer   'character map
      Dim nKeys(16) As Integer     'xor keys used for pArray(x) xor nkeys(x)
      Dim s0(512) As Integer       'swap-box data used to map character table
      Dim nArray(16) As Integer    'name array data
      Dim pArray(16) As Integer    'password array data
      Dim n As Integer             'for next loop counter
      Dim nPtr As Integer          'name pointer (used for counting)
      Dim cPtr As Integer          'character pointer (used for counting)
      Dim cFlip As Boolean         'character flip (used to flip between numeric and alpha)
      Dim sIni As Integer          'holds s-box values
      Dim temp As Integer          'holds s-box values
      Dim rtn As Integer           'holds generated key values used agains chr map
      Dim gKey As String           'generated key as string
      Dim nLen As Integer          'number of chr's in name
      Dim pLen As Integer          'number of chr's in password
      Dim kPtr As Integer          'key pointer
      Dim sPtr As Integer          'space pointer (used in hex key)
      Dim nOffset As Integer       'name offset
      Dim pOffset As Integer       'password offset
      Dim tOffset As Integer       'total offset
      Dim KeySize As Integer       'the size of the key to make

      Const nXor As Integer = 18   'name xor value
      Const pXor As Integer = 25   'password xor value
      Const cLw As Integer = 65    'character lower limit 65 = A ** do not change **
      Const nLw As Integer = 48    'number lower limit 48 = 0 ** do not change **
      Const sOffset As Integer = 0 'character map offset
      '****************************************************************************
      'Thanks to Chris Fournier for his suggestions for adding support for        *
      'Strings, Objects and String() as arrays                                    *
      'Your comments please                                                       *
      '****************************************************************************
      Dim VarType As String
      Dim kName As String
      Dim AryCtl As Integer
      Dim AryCtrl As Control

20        VarType = TypeName(kNamev)
30        Select Case VarType
              Case "String"
40                kName = kNamev
50            Case "TextBox"
60                kName = kNamev.Text
70            Case "Object"
80                For Each AryCtrl In kNamev
90                    If AryCtrl.Text <> "" Then
100                       kName = kName & AryCtrl.Text & "|"
110                   End If
120               Next
130               kName = Left(kName, Len(kName) - 1)
140           Case "String()"
150               For AryCtl = LBound(kNamev) To UBound(kNamev)
160                   If kNamev(AryCtl) <> "" Then
170                       kName = kName & kNamev(AryCtl) & "|"
180                   End If
190               Next
200               kName = Left(kName, Len(kName) - 1)
210               Case Else
220                   MsgBox VarType & " is an unsupported type to be passed to KeyGen"
230       End Select
      '****************************************************************************
          
240       nLen = Len(kName)
250       pLen = Len(kPass)
          
          'password xor keys ** change to make keygen unique **
260       nKeys(1) = 46
270       nKeys(2) = 89
280       nKeys(3) = 142
290       nKeys(4) = 63
300       nKeys(5) = 231
310       nKeys(6) = 32
320       nKeys(7) = 129
330       nKeys(8) = 51
340       nKeys(9) = 28
350       nKeys(10) = 97
360       nKeys(11) = 248
370       nKeys(12) = 41
380       nKeys(13) = 136
390       nKeys(14) = 53
400       nKeys(15) = 78
410       nKeys(16) = 164
          
420       sIni = 0
          
          'set s boxes
430       For n = 0 To 512
440           s0(n) = n
450       Next n
          
460       For n = 0 To 512
470           sIni = (sOffset + sIni + n) Mod 256
480           temp = s0(n)
490           s0(n) = s0(sIni)
500           s0(sIni) = temp
510       Next n
          
520       If kType = 1 Then       '(numeric)
              
530           nPtr = 0
540           KeySize = 16
550           gKey = String(16, " ")
              
560           For n = 0 To 512
570               cTable(s0(n)) = (nLw + (nPtr))
580               nPtr = nPtr + 1
590               If nPtr = 10 Then nPtr = 0
600           Next n
          
610       ElseIf kType = 2 Then   '(alphanumeric)
              
620           nPtr = 0
630           cPtr = 0
640           KeySize = 16
650           gKey = String(16, " ")
              
660           cFlip = False
670           For n = 0 To 512
680               If cFlip Then
690                   cTable(s0(n)) = (nLw + nPtr)
700                   nPtr = nPtr + 1
710                   If nPtr = 10 Then nPtr = 0
720                   cFlip = False
730               Else
740                   cTable(s0(n)) = (cLw + cPtr)
750                   cPtr = cPtr + 1
760                   If cPtr = 26 Then cPtr = 0
770                   cFlip = True
780               End If
790           Next n
          
800       Else  '(hex)
          
810           KeySize = 8
820           gKey = String(19, " ")
              
830       End If
          
840       kPtr = 1
          
850       For n = 1 To nLen 'name
860         nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
870         nOffset = nOffset + nArray(kPtr)
880         kPtr = kPtr + 1
890           If kPtr = 9 Then kPtr = 1
900       Next n
          
910       For n = 1 To pLen 'password
920         pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
930         pOffset = pOffset + pArray(kPtr)
940         kPtr = kPtr + 1
950           If kPtr = 9 Then kPtr = 1
960       Next n
          
970       tOffset = (nOffset + pOffset) Mod 512
          
980       kPtr = 1
990       sPtr = 1
1000      For n = 1 To KeySize
1010        pArray(n) = pArray(n) Xor nKeys(n)
1020        rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
            
1030        If kType = 3 Then 'hex key
1040              If rtn < 16 Then
1050                  Mid(gKey, kPtr, 2) = "0" & Hex(rtn)
1060              Else
1070                  Mid(gKey, kPtr, 2) = Hex(rtn)
1080              End If
1090                  If sPtr = 2 And kPtr < 18 Then
1100                      kPtr = kPtr + 1
1110                      Mid(gKey, kPtr + 1, 1) = "-"
1120                  End If
1130              kPtr = kPtr + 2
1140              sPtr = sPtr + 1
1150              If sPtr = 3 Then sPtr = 1
1160        Else  'numeric - alphanumeric key
1170          Mid(gKey, n, 1) = Chr(cTable(rtn))
1180        End If
1190      Next
1200      KeyGen = gKey
End Function
Function Equals(chr1 As String, chr2 As String, chr3 As String, chr4 As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : Equals
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : Comparación
'---------------------------------------------------------------------------------------

10       On Error GoTo Equals_Error

20        Equals = False
          
30        Select Case chr1
              Case chr2, chr3, chr4
40                Equals = True
50        End Select
60        Select Case chr2
              Case chr3, chr4
70                Equals = True
80        End Select
90        If chr3 = chr4 Then
100           Equals = True
110       End If

120      On Error GoTo 0
130      Exit Function

Equals_Error:
140       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure Equals of Módulo modKeyGen in Line " & Erl& & vbNewLine)
End Function

'****************************************************************************
Sub CheckRegister()
'---------------------------------------------------------------------------------------
' Procedure : CheckRegister
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   :
'---------------------------------------------------------------------------------------
'Javier Orellana 25/09/2006 23:00 Hrs.
'Modificado para Password Eliminator .XLS 23/01/2007 21:37 Hrs.
'Chequea si el software ha sido registrado
'Este Algoritmo Tambien se Repite en el Formulario de Registro  no permite que las claves de registro
'finalizen con 4 letras iguales y ademas la ultima parte de la clave de registro
'debe comenzar con la Letra A
Dim strAuxReg As String
Dim strAux As String
Dim strReg As String
          
10       On Error GoTo CheckRegister_Error

20        blnRegistrado = False
          
30        strAuxReg = ReadINI(App.Path & "\config.cfg", "Registro", "Clave")
40        strAux = Right(strAuxReg, 4)
50        If Not Equals(Left(strAux, 1), Mid(strAux, 2, 1), Mid(strAux, 3, 1), Right(strAux, 1)) Then
60            If Left(strAux, 1) = "A" Then
70                strReg = KeyGen(Left(strAux, 2), Right(strAux, 2), 3) & "-" & strAux
80                If strReg = strAuxReg Then
90                    blnRegistrado = True
100               End If
110           End If
120       End If
          
130       If Not blnRegistrado Then
140           frmRegistro.Show vbModal
150       End If

160      On Error GoTo 0
170      Exit Sub

CheckRegister_Error:
180       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure CheckRegister of Módulo modKeyGen in Line " & Erl& & vbNewLine)
End Sub

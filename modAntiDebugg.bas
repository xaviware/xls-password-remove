Attribute VB_Name = "modAntiDebugg"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : modAntiDebugg
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : ' Name : Detect Numega Smart Check, W32Dasm, Hex Workshop
'---------------------------------------------------------------------------------------

'*-------------------------------------
'Stores the truncated class names of
'applications which are to be monitored in the
'system while this application is running
Public Sub FillClassName()
'---------------------------------------------------------------------------------------
' Procedure : FillClassName
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   :
'---------------------------------------------------------------------------------------
Dim i As Integer
          
10       On Error GoTo FillClassName_Error

          'Clear the Array
20        For i = 0 To NoOfAppClassMonitored
30            AppClassName(i) = ""
40        Next i
50        AppClassName(0) = "HexWorks" 'Hex Workshop 3.1
60        AppClassName(1) = "OWL_Window" 'WinDasm
70        AppClassName(2) = "NMSCMW" 'for various versions of Numega Smart Check
80        AppClassName(3) = "UltraEdit-32" 'Ultra Edit
          
          'WinAmp Playlist, WinAmp Mnibrowser, WinAmp version x.x, Winamp Equaliser
          'AppClassName(4) = "Notepad" 'Notepad
          'AppClassName(5) = "UltraEdit-32" 'Ultra Edit

90       On Error GoTo 0
100      Exit Sub

FillClassName_Error:
110       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure FillClassName of Módulo modAntiDebugg in Line " & Erl& & vbNewLine)
End Sub
Public Function AppPresent(TgtClassName As String, frmName As Object) As Long
'---------------------------------------------------------------------------------------
' Procedure : AppPresent
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   :
'---------------------------------------------------------------------------------------
'Check if the windows classname matches with
'our list in the array (TgtClassName i.e. target
'classname)

10       On Error GoTo AppPresent_Error

20        AppPresent = SearchAppsByClassName(TgtClassName, frmName)

30       On Error GoTo 0
40       Exit Function

AppPresent_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure AppPresent of Módulo modAntiDebugg in Line " & Erl& & vbNewLine)
End Function
Function SearchAppsByClassName(TgtClassName As String, frmName As Object)
'---------------------------------------------------------------------------------------
' Procedure : SearchAppsByClassName
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   :
'---------------------------------------------------------------------------------------
'Searches all the applications running in
'Windows environment either in the
'foreground or in the background
Dim ThisAppHandle As Long
Dim NextAppHandle As Long
Dim AppClassName As String
          
10       On Error GoTo SearchAppsByClassName_Error

          'Get your current application owner's
          'handle
20        ThisAppHandle = getwindow(frmName.hwnd, GW_OWNER)
          'Pass it on as a seed for the next handle
          'so all the windows in the z order can be
          'searched
30        NextAppHandle = ThisAppHandle
          'Perform the iterations in the Z order till
          'all the running applications (either foreground
          'or background) have been run through
40        Do While NextAppHandle <> 0
50            DoEvents
              'Get the handle of the next window
              'in the Z order
60            NextAppHandle = getwindow(NextAppHandle, GW_HWNDNEXT)
              'Retrieve its window's classname
70            AppClassName = GetAppClassName(NextAppHandle)
              'If a part of TgtClassName is found
              'then retrieve this applications handle
80            If InStr(1, LCase(Trim(AppClassName)), LCase(Trim(TgtClassName))) <> 0 Then
90                SearchAppsByClassName = NextAppHandle
100               Exit Do
110               Exit Function
120           End If
130       Loop

140      On Error GoTo 0
150      Exit Function

SearchAppsByClassName_Error:
160       WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure SearchAppsByClassName of Módulo modAntiDebugg in Line " & Erl& & vbNewLine)
End Function
Private Function GetAppClassName(HwndWindow As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : GetAppClassName
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Get the Window's Class Name
'---------------------------------------------------------------------------------------
Dim BufferAll&
Dim WindowClassName$

10       On Error GoTo GetAppClassName_Error

20        WindowClassName$ = String(100, Chr(0))
30        BufferAll& = GetClassName(HwndWindow, WindowClassName$, 100)
40        GetAppClassName = Left(WindowClassName$, BufferAll&)
50        DoEvents

60       On Error GoTo 0
70       Exit Function

GetAppClassName_Error:
80        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure GetAppClassName of Módulo modAntiDebugg in Line " & Erl& & vbNewLine)
End Function
Public Sub KillWin(hwnd As Long)
'---------------------------------------------------------------------------------------
' Procedure : KillWin
' Author    : Administrador - JAOO
' Date      : 20/01/2008
' Purpose   : 'Close the selected application
'---------------------------------------------------------------------------------------
Dim dummy

10       On Error GoTo KillWin_Error

20        dummy = sendmessagebystring(hwnd, WM_CLOSE, 0, 0)
          'If frmcheck.chkFound = 1 Then
              'Raise an exception
          '    RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
          'End If

30       On Error GoTo 0
40       Exit Sub

KillWin_Error:
50        WriteLogToFile (Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error " & Err.Number & " (" & Err.Description & ") in procedure KillWin of Módulo modAntiDebugg in Line " & Erl& & vbNewLine)
End Sub

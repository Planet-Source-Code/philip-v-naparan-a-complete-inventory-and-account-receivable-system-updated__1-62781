Attribute VB_Name = "modAPI"
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

Option Explicit

'===============================
'API Declarations and Constant
'===============================

'For tracking mouse cursor position
Public Declare Function GetCursorPos Lib "user32" _
            (lpPoint As POINTAPI) As Long
            
Public Type POINTAPI
        x As Long
        y As Long
End Type


'For memory status
Public Declare Sub GlobalMemoryStatus Lib "kernel32" _
                (lpBuffer As MEMORYSTATUS)
                
Private Type MEMORYSTATUS 'Type variable for memory info
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


'Use to get the window top,left,right and buttom position
Public Declare Function GetWindowRect Lib "user32" _
                (ByVal hwnd As Long, _
                lpRect As RECT) As Long
                
Public Type RECT 'Type variable for window rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
                
Public Const WM_CLOSE = &H10 'Message use to perform close
Public Const WM_ACTIVATE = &H6 'Message use to perform activate



'Use to set the parent
Public Declare Function SetParent Lib "user32" _
                (ByVal hWndChild As Long, _
                ByVal hWndNewParent As Long) As Long

'Use for setting windows on top
Public Declare Function SetWindowPos Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'API for opening a browser
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long


Public MEM_STAT As MEMORYSTATUS
'API used to change the form border
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Related contstant (see API used to change the form border)
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000
Public Const WS_DLGFRAME = &H400000


Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4
Public Const SWPFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

'Function used to change the form border (see API used to change the form border)
Public Function FormStyle(ByRef srcHWND As Long, ByVal Bit As Long, ByVal Value As Boolean) As Boolean
    Dim nStyle As Long

    nStyle = GetWindowLong(srcHWND, GWL_STYLE)

    If Value Then
        nStyle = nStyle Or Bit
    Else
        nStyle = nStyle And Not Bit
    End If
    Call SetWindowLong(srcHWND, GWL_STYLE, nStyle)
    SetWindowPos srcHWND, 0, 0, 0, 0, 0, SWPFlags

    FormStyle = (nStyle = GetWindowLong(srcHWND, GWL_STYLE))
End Function

Public Sub OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub

Public Function OnTop(ByRef hwnd As Long, Optional BringOnTop As Boolean = True) As Long
    If BringOnTop = True Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Function





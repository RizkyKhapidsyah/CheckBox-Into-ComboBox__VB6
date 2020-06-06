Attribute VB_Name = "Module1"
Public Const EC_LEFTMARGIN = &H1
Public Const EC_RIGHTMARGIN = &H2
Public Const EC_USEFONTINFO = &HFFFF&
Public Const EM_SETMARGINS = &HD3&
Public Const EM_GETMARGINS = &HD4&
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, _
ByVal lpszWindow As String) As Long
Declare Function SendMessageLong Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
As Long, ByVal wParam As Long, ByVal lParam As Long) _
As Long






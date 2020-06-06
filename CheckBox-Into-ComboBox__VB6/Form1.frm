VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menambahkan CheckBox ke ComboBox"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddCheckToCombo(ByRef chkThis _
As CheckBox, ByRef cboThis As ComboBox)
Dim lhWnd As Long
Dim lMargin As Long
lhWnd = FindWindowEx(cboThis.hwnd, 0, "EDIT", vbNullString)
  If (lhWnd <> 0) Then
     lMargin = chkThis.Width \ Screen.TwipsPerPixelX _
               + 2
     SendMessageLong lhWnd, EM_SETMARGINS, _
     EC_LEFTMARGIN, lMargin
     chkThis.BackColor = cboThis.BackColor
     chkThis.Move cboThis.Left + 3 * _
     Screen.TwipsPerPixelX, cboThis.Top + 2 * _
     Screen.TwipsPerPixelY, _
     chkThis.Width, cboThis.Height - 4 * _
     Screen.TwipsPerPixelY
     chkThis.ZOrder
  End If
End Sub

Private Sub Form_Load()
   AddCheckToCombo Check1, Combo1
End Sub



Attribute VB_Name = "MdlDesign"

'move form
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'-------------
'make transparent
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

Private Declare Function UpdateLayeredWindow Lib "user32" _
(ByVal hwnd As Long, ByVal hDCDst As Long, pptDst As Any, _
psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, _
ByVal pblend As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
MakeTransparent = 1
Else
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
Msg = Msg Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
MakeTransparent = 0
End If
If Err Then
MakeTransparent = 2
End If
End Function

Sub MaxLength_8()
    FrmInput.TxtPenGapok.MaxLength = 8
    FrmInput.TxtPenMakan.MaxLength = 8
    FrmInput.TxtPenTransport.MaxLength = 8
    FrmInput.TxtPenLembur.MaxLength = 8
    FrmInput.TxtPenInsHarian.MaxLength = 8
    FrmInput.TxtPenIns.MaxLength = 8
    FrmInput.TxtPenJHT.MaxLength = 8
    FrmInput.TxtPenJKN.MaxLength = 8
    FrmInput.TxtPenPensiun.MaxLength = 8
    FrmInput.TxtPenPajak.MaxLength = 8
    FrmInput.TxtPenLain.MaxLength = 8
    
    FrmInput.TxtPotMakan.MaxLength = 8
    FrmInput.TxtPotJHT.MaxLength = 8
    FrmInput.TxtPotJKN.MaxLength = 8
    FrmInput.TxtPotPensiun.MaxLength = 8
    FrmInput.TxtPotAbsen.MaxLength = 8
    FrmInput.TxtPotPajak.MaxLength = 8
    FrmInput.TxtPotLain.MaxLength = 8
End Sub

Sub MaxLength_13()
    FrmInput.TxtPenGapok.MaxLength = 13
    FrmInput.TxtPenMakan.MaxLength = 13
    FrmInput.TxtPenTransport.MaxLength = 13
    FrmInput.TxtPenLembur.MaxLength = 13
    FrmInput.TxtPenInsHarian.MaxLength = 13
    FrmInput.TxtPenIns.MaxLength = 13
    FrmInput.TxtPenJHT.MaxLength = 13
    FrmInput.TxtPenJKN.MaxLength = 13
    FrmInput.TxtPenPensiun.MaxLength = 13
    FrmInput.TxtPenPajak.MaxLength = 13
    FrmInput.TxtPenLain.MaxLength = 13
    
    FrmInput.TxtPotMakan.MaxLength = 13
    FrmInput.TxtPotJHT.MaxLength = 13
    FrmInput.TxtPotJKN.MaxLength = 13
    FrmInput.TxtPotPensiun.MaxLength = 13
    FrmInput.TxtPotAbsen.MaxLength = 13
    FrmInput.TxtPotPajak.MaxLength = 13
    FrmInput.TxtPotLain.MaxLength = 13
End Sub


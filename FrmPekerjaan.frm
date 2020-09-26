VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPekerjaan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPekerjaan.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Masa Kerja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5040
      TabIndex        =   27
      Top             =   2880
      Width           =   2295
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   435
         TabIndex        =   29
         Top             =   720
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   435
         TabIndex        =   28
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(Bulan)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   31
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(Tahun)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   30
         Top             =   360
         Width           =   900
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   1065
         Picture         =   "FrmPekerjaan.frx":DBC9
         Top             =   720
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   8
         Left            =   240
         Picture         =   "FrmPekerjaan.frx":10DD2
         Top             =   720
         Width           =   195
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   1065
         Picture         =   "FrmPekerjaan.frx":13FD1
         Top             =   360
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   7
         Left            =   240
         Picture         =   "FrmPekerjaan.frx":171DA
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.ListBox LstUrut 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "FrmPekerjaan.frx":1A3D9
      Left            =   2670
      List            =   "FrmPekerjaan.frx":1A3DB
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker DTKeluar 
      Height          =   300
      Left            =   2685
      TabIndex        =   25
      Top             =   3840
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116326403
      CurrentDate     =   43796
   End
   Begin VB.TextBox TxtAlamat 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   3090
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   5760
      Width           =   3090
   End
   Begin MSComCtl2.DTPicker DTMasuk 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   2685
      TabIndex        =   21
      Top             =   3480
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   116326403
      CurrentDate     =   43786
   End
   Begin VB.TextBox TxtJabatan 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2835
      TabIndex        =   20
      Top             =   3120
      Width           =   1840
   End
   Begin VB.TextBox TxtBagian 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2835
      TabIndex        =   19
      Top             =   2760
      Width           =   1840
   End
   Begin VB.TextBox TxtPerusahaan 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2835
      TabIndex        =   18
      Top             =   2400
      Width           =   3405
   End
   Begin VB.TextBox TxtKode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2835
      TabIndex        =   17
      Top             =   2040
      Width           =   1840
   End
   Begin VB.TextBox TxtNIK 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2835
      TabIndex        =   16
      Top             =   1680
      Width           =   1840
   End
   Begin VB.CheckBox ChcNPWP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "NPWP"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   15
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   3195
   End
   Begin VB.CheckBox ChcJKN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "BPJS Kesehatan (JKN)"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   14
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   3195
   End
   Begin VB.CheckBox ChcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   7
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   3195
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   7440
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox ChcJHT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "BPJS Ketenagakerjaan (JHT)"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   5
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   3195
   End
   Begin VB.TextBox TxtNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   2835
      TabIndex        =   1
      Text            =   "Auto.."
      Top             =   1260
      Width           =   640
   End
   Begin VB.Timer Timer2 
      Left            =   8040
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   360
   End
   Begin VB.Label LblTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   "dateeeeeee..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   7440
      Width           =   3420
   End
   Begin VB.Image Image8 
      Height          =   750
      Left            =   6120
      Picture         =   "FrmPekerjaan.frx":1A3DD
      Top             =   5760
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   6
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":1A75D
      Top             =   5760
      Width           =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   22
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   5
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":1AAE2
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   4680
      Picture         =   "FrmPekerjaan.frx":1DCE1
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   4
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":20EEA
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   4680
      Picture         =   "FrmPekerjaan.frx":240E9
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   3
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":272F2
      Top             =   2400
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   6240
      Picture         =   "FrmPekerjaan.frx":2A4F1
      Top             =   2400
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":2D6FA
      Top             =   2040
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   4680
      Picture         =   "FrmPekerjaan.frx":308F9
      Top             =   2040
      Width           =   195
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   600
      Y1              =   1080
      Y2              =   6720
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   1080
      Y2              =   6720
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":33B02
      Top             =   1680
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   4680
      Picture         =   "FrmPekerjaan.frx":36D01
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Keluar"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   13
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Masuk"
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   12
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jabatan"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   11
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bagian"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Perusahaan"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   9
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Perusahaan"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   1860
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3480
      Picture         =   "FrmPekerjaan.frx":39F0A
      Top             =   1260
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   2640
      Picture         =   "FrmPekerjaan.frx":3D21D
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NIK"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ".:: History Pekerjaan ::."
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   15
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Close"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6885
      TabIndex        =   2
      Top             =   6960
      Width           =   1590
   End
   Begin VB.Image ImgExit 
      Height          =   555
      Left            =   6795
      Picture         =   "FrmPekerjaan.frx":4041C
      Tag             =   "1"
      Top             =   6795
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Urut"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1305
      Width           =   780
   End
End
Attribute VB_Name = "FrmPekerjaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim JmlPersen As Integer

Private Sub ChcStatus_Click()
    If ChcStatus.Value = Checked Then
        DTKeluar.Value = Now
        DTKeluar.Enabled = True
    ElseIf ChcStatus.Value = Unchecked Then
        DTKeluar.Value = "01/01/1900"
        DTKeluar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    MakeTransparent Me.hwnd, 0
    JmlPersen = 0
    Me.Timer1.Interval = 50
    
    DTMasuk.CustomFormat = "dd MMM yyyy"
    DTKeluar.CustomFormat = "dd MMM yyyy"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ImgExit.Tag = 2 Then
        Me.ImgExit.Picture = FrmGaji.ImgLst32.ListImages(8).Picture
        Me.ImgExit.Tag = 1
    End If
End Sub

Private Sub Image3_Click()
'    Me.List2.Visible = True
'    Me.List2.SetFocus
End Sub

Private Sub ImgExit_Click()
    Me.Timer1.Interval = 0
    Me.Timer2.Interval = 50
End Sub


Private Sub Label2_Click()
    ImgExit_Click
    FrmInput.Timer5.Enabled = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ImgExit.Tag = 1 Then
        Me.ImgExit.Picture = FrmGaji.ImgLst32.ListImages(9).Picture
        Me.ImgExit.Tag = 2
    End If
End Sub

Private Sub List2_Click()
'    Me.Text3.Text = Me.List2.Text
'    Me.List2.Visible = False
End Sub

Private Sub Timer1_Timer()
    If JmlPersen < 230 Then
        JmlPersen = JmlPersen + 10
    Else
        JmlPersen = 230
        Me.Timer1.Interval = 0
    End If
    MakeTransparent Me.hwnd, JmlPersen
End Sub

Private Sub Timer2_Timer()
    If JmlPersen > 0 Then
        JmlPersen = JmlPersen - 10
    Else
        Unload Me
        Exit Sub
    End If
    MakeTransparent Me.hwnd, JmlPersen
End Sub

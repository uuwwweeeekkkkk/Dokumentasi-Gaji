VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Gaji Juan Selama Bekerja"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
   Icon            =   "FrmGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FrmGaji.frx":2832
   MousePointer    =   99  'Custom
   ScaleHeight     =   8100
   ScaleWidth      =   13260
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImgLst32 
      Left            =   -200
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":2C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":30C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":3DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":420E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":7395
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":B53B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":F32A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Masukan Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Picture         =   "FrmGaji.frx":130C5
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   5640
      Picture         =   "FrmGaji.frx":133CF
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   3720
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame FrmCari 
      Caption         =   "*Tekan Enter Untuk Mencari Data :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   2760
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   7575
      Begin VB.ComboBox CboCari 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmGaji.frx":1403D
         Left            =   480
         List            =   "FrmGaji.frx":1403F
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtCari 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MousePointer    =   3  'I-Beam
         TabIndex        =   19
         Top             =   480
         Width           =   4095
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2880
         Y1              =   650
         Y2              =   650
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.CommandButton CmdCetak 
      Caption         =   "&Cetak ke Excel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      Picture         =   "FrmGaji.frx":14041
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Total Perhitungan :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   12495
      Begin VB.TextBox TxtTotGaji 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtTotPen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5160
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtTotPot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9600
         TabIndex        =   12
         Text            =   "0"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label LblTotGaji 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Total Gaji Dibayar"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label LblTotPen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Total Pendapatan"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label LblTotPot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Total Potongan"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9600
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      Picture         =   "FrmGaji.frx":14483
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1695
   End
   Begin VB.ComboBox CboBahasa 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      Picture         =   "FrmGaji.frx":148C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgLst16 
      Left            =   -200
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":15737
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":164C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":167DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGaji.frx":17565
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Bersihkan List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      Picture         =   "FrmGaji.frx":182EF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame FrmSecurity 
      Caption         =   "*Tekan Enter Kata Sandi Anda Untuk Menampilkan Data :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1300
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CheckBox ChcPass 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4850
         TabIndex        =   23
         Top             =   620
         Width           =   255
      End
      Begin VB.TextBox TxtPass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         IMEMode         =   3  'DISABLE
         Left            =   720
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Picture         =   "FrmGaji.frx":18731
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton CmdTampilkan 
      Caption         =   "&Tampilkan Data"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Picture         =   "FrmGaji.frx":18B73
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4755
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8387
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   7605
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "FrmGaji.frx":18FB5
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3792
            MinWidth        =   3792
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "FrmGaji.frx":19407
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14340
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3369
            MinWidth        =   3369
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Visible         =   0   'False
      Width           =   12495
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   12840
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "FrmGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sHari As String
Dim aHari
Dim uHari
Dim Kesempatan As Integer
Dim GntGmbr As Boolean
Dim GntIcon As Byte

Private Sub CboBahasa_Click()
    Picture1.Visible = True
    If CboBahasa.Text = "English" Then
        Frame1.Caption = "&Total Calculation : "
        LblTotGaji.Caption = "&Total Salary Paid"
        LblTotPen.Caption = "&Total Income"
        LblTotPot.Caption = "&Total Pieces"
        CmdInput.Caption = "&Input Data"
        CmdTampilkan.Caption = "&Show Data"
        CmdBatal.Caption = "&Cancel"
        CmdCari.Caption = "&Search Data"
        CmdHapus.Caption = "&Clear List"
        CmdCetak.Caption = "&Print to Excel"
        CmdKeluar.Caption = "&Exit"
        FrmSecurity.Caption = "*Press Enter Your Password For Display Data :"
        FrmCari.Caption = "*Press Enter For Searching Data :"
        CboCari.Clear
        CboCari.AddItem "Month"
        CboCari.AddItem "Year"
        CboCari.AddItem "Position"
        CboCari.AddItem "Company"
        CboCari.Text = "Month"
        Call Load_Waktu
        Call Clm_Eng
    ElseIf CboBahasa.Text = "Indonesia" Then
        Frame1.Caption = "&Total Perhitungan : "
        LblTotGaji.Caption = "&Total Gaji Dibayar"
        LblTotPen.Caption = "&Total Pendapatan"
        LblTotPot.Caption = "&Total Potongan"
        CmdInput.Caption = "&Masukan Data"
        CmdTampilkan.Caption = "&Tampilkan Data"
        CmdBatal.Caption = "&Batal"
        CmdCari.Caption = "&Cari Data"
        CmdHapus.Caption = "&Bersihkan List"
        CmdCetak.Caption = "&Cetak ke Excel"
        CmdKeluar.Caption = "&Keluar"
        FrmSecurity.Caption = "*Tekan Enter Kata Sandi Anda Untuk Menampilkan Data :"
        FrmCari.Caption = "*Tekan Enter Untuk Mencari Data :"
        CboCari.Clear
        CboCari.AddItem "Bulan"
        CboCari.AddItem "Tahun"
        CboCari.AddItem "Jabatan"
        CboCari.AddItem "Perusahaan"
        CboCari.Text = "Bulan"
        Call Load_Waktu
        Call Clm_Ind
    End If
    Picture1.Visible = False
End Sub

Sub Clm_Eng()
    ListView1.ColumnHeaders.Clear
    With ListView1.ColumnHeaders
        .Add , , "No", Width / 16.9, lvwColumnLeft, 1
        .Add , , "Month", Width / 10, lvwColumnLeft, 1
        .Add , , "Year", Width / 11, lvwColumnLeft, 1
        .Add , , "Company", Width / 4.8, lvwColumnLeft, 1
        .Add , , "Position", Width / 8, lvwColumnLeft, 1
        .Add , , "Paid Salary", Width / 6.7, lvwColumnRight, 3
        .Add , , "Total Income", Width / 5.4, lvwColumnRight, 3
        .Add , , "Total Pieces", Width / 5.9, lvwColumnRight, 4
        .Add , , "Basic Salary", Width / 7.6, lvwColumnRight, 3
        .Add , , "Food Money", Width / 7.6, lvwColumnRight, 3
        .Add , , "Transportation Money", Width / 5.9, lvwColumnRight, 3
        .Add , , "Overtime Pay", Width / 7.1, lvwColumnRight, 3
        .Add , , "Daily Incentive", Width / 5.6, lvwColumnRight, 3
        .Add , , "Incentive", Width / 8.8, lvwColumnRight, 3
        .Add , , "JHT", Width / 8.8, lvwColumnRight, 3
        .Add , , "BPJS Health", Width / 5.9, lvwColumnRight, 3
        .Add , , "Pension Insurance", Width / 5.6, lvwColumnRight, 3
        .Add , , "PPh21 Tax", Width / 7.1, lvwColumnRight, 3
        .Add , , "Etc", Width / 8.2, lvwColumnRight, 3
        .Add , , "Food Money", Width / 7.6, lvwColumnRight, 4
        .Add , , "Abcent", Width / 8.8, lvwColumnRight, 4
        .Add , , "JHT", Width / 8.8, lvwColumnRight, 4
        .Add , , "BPJS Health", Width / 5.9, lvwColumnRight, 4
        .Add , , "Pension Insurance", Width / 5.6, lvwColumnRight, 4
        .Add , , "PPh21 Tax", Width / 7.1, lvwColumnRight, 4
        .Add , , "Etc", Width / 8.2, lvwColumnRight, 4
        .Add , , "Last Update", Width / 6, lvwColumnLeft, 1
    End With
End Sub

Sub Clm_Ind()
    ListView1.ColumnHeaders.Clear
    With ListView1.ColumnHeaders
        .Add , , "No", Width / 16.9, lvwColumnLeft, 1
        .Add , , "Bulan", Width / 10, lvwColumnLeft, 1
        .Add , , "Tahun", Width / 11, lvwColumnLeft, 1
        .Add , , "Perusahaan", Width / 4.8, lvwColumnLeft, 1
        .Add , , "Jabatan", Width / 8, lvwColumnLeft, 1
        .Add , , "Gaji Dibayar", Width / 6.7, lvwColumnRight, 3
        .Add , , "Total Pendapatan", Width / 5.4, lvwColumnRight, 3
        .Add , , "Total Potongan", Width / 5.9, lvwColumnRight, 4
        .Add , , "Gaji Pokok", Width / 7.6, lvwColumnRight, 3
        .Add , , "Uang Makan", Width / 7.6, lvwColumnRight, 3
        .Add , , "Uang Transport", Width / 5.9, lvwColumnRight, 3
        .Add , , "Uang Lembur", Width / 7.1, lvwColumnRight, 3
        .Add , , "Insentif Harian", Width / 5.6, lvwColumnRight, 3
        .Add , , "Insentif", Width / 8.8, lvwColumnRight, 3
        .Add , , "JHT", Width / 8.8, lvwColumnRight, 3
        .Add , , "BPJS Kesehatan", Width / 5.9, lvwColumnRight, 3
        .Add , , "Jaminan Pensiun", Width / 5.6, lvwColumnRight, 3
        .Add , , "Pajak PPh21", Width / 7.1, lvwColumnRight, 3
        .Add , , "Lain-Lain", Width / 8.2, lvwColumnRight, 3
        .Add , , "Uang Makan", Width / 7.6, lvwColumnRight, 4
        .Add , , "Absen", Width / 8.8, lvwColumnRight, 4
        .Add , , "JHT", Width / 8.8, lvwColumnRight, 4
        .Add , , "BPJS Kesehatan", Width / 5.9, lvwColumnRight, 4
        .Add , , "Jaminan Pensiun", Width / 5.6, lvwColumnRight, 4
        .Add , , "Pajak PPh21", Width / 7.1, lvwColumnRight, 4
        .Add , , "Lain-Lain", Width / 8.2, lvwColumnRight, 4
        .Add , , "Pembaharuan Terakhir", Width / 6, lvwColumnLeft, 1
    End With
End Sub

Private Sub CboCari_Click()
On Error Resume Next
    TxtCari.SetFocus
    TxtCari.Text = ""
End Sub

Private Sub CmdBatal_Click()
    If FrmSecurity.Visible = True Then
        TxtPass.Text = ""
        ChcPass.Value = False
        FrmSecurity.Visible = False
        CmdBatal.Enabled = False
        CmdTampilkan.Enabled = True
    ElseIf FrmCari.Visible = True Then
        Call Bersih_Cari
    End If
End Sub

Private Sub CmdCari_Click()
    FrmCari.Visible = True
    TxtCari.SetFocus
    TxtCari.MaxLength = "30"
    CmdBatal.Enabled = True
    
    If CboBahasa.Text = "Indonesia" Then
        CboCari.Clear
        CboCari.AddItem "Bulan"
        CboCari.AddItem "Tahun"
        CboCari.AddItem "Jabatan"
        CboCari.AddItem "Perusahaan"
        CboCari.Text = "Bulan"
    ElseIf CboBahasa.Text = "English" Then
        CboCari.Clear
        CboCari.AddItem "Month"
        CboCari.AddItem "Year"
        CboCari.AddItem "Position"
        CboCari.AddItem "Company"
        CboCari.Text = "Month"
    End If
End Sub

Private Sub ChcPass_Click()
    If ChcPass.Value = 1 Then
        TxtPass.PasswordChar = ""
    ElseIf ChcPass.Value = 0 Then
        TxtPass.PasswordChar = "*"
    End If
    TxtPass.SetFocus
End Sub

Private Sub CmdCetak_Click()
    On Error Resume Next
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim I As Integer
    
    Picture1.Visible = True
    Me.MousePointer = vbHourglass
        Set ExcelObj = CreateObject("Excel.Application")
            Set ExcelBook = ExcelObj.Workbooks.Add
                Set ExcelSheet = ExcelBook.Worksheets(1)

        ExcelObj.Visible = True
        With ExcelSheet
            For I = 1 To ListView1.ListItems.Count
                .Cells(I, 1) = ListView1.ListItems(I).Text
                .Cells(I, 2) = ListView1.ListItems(I).SubItems(1)
                .Cells(I, 3) = ListView1.ListItems(I).SubItems(2)
                .Cells(I, 4) = ListView1.ListItems(I).SubItems(3)
                .Cells(I, 5) = ListView1.ListItems(I).SubItems(4)
                .Cells(I, 6) = ListView1.ListItems(I).SubItems(5)
                .Cells(I, 7) = ListView1.ListItems(I).SubItems(6)
                .Cells(I, 8) = ListView1.ListItems(I).SubItems(7)
                .Cells(I, 9) = ListView1.ListItems(I).SubItems(8)
                .Cells(I, 10) = ListView1.ListItems(I).SubItems(9)
                .Cells(I, 11) = ListView1.ListItems(I).SubItems(10)
                .Cells(I, 12) = ListView1.ListItems(I).SubItems(11)
                .Cells(I, 13) = ListView1.ListItems(I).SubItems(12)
                .Cells(I, 14) = ListView1.ListItems(I).SubItems(13)
                .Cells(I, 15) = ListView1.ListItems(I).SubItems(14)
                .Cells(I, 16) = ListView1.ListItems(I).SubItems(15)
                .Cells(I, 17) = ListView1.ListItems(I).SubItems(16)
                .Cells(I, 18) = ListView1.ListItems(I).SubItems(17)
                .Cells(I, 19) = ListView1.ListItems(I).SubItems(18)
                .Cells(I, 20) = ListView1.ListItems(I).SubItems(19)
                .Cells(I, 21) = ListView1.ListItems(I).SubItems(20)
                .Cells(I, 22) = ListView1.ListItems(I).SubItems(21)
                .Cells(I, 23) = ListView1.ListItems(I).SubItems(22)
                .Cells(I, 24) = ListView1.ListItems(I).SubItems(23)
                .Cells(I, 25) = ListView1.ListItems(I).SubItems(24)
                .Cells(I, 26) = ListView1.ListItems(I).SubItems(25)
                .Cells(I, 27) = ListView1.ListItems(I).SubItems(26)
            Next
                .Cells(I, 3) = "Total Gaji Dibayar : " & TxtTotGaji.Text
                .Cells(I, 6) = "Total Pedapatan : " & TxtTotPen.Text
                .Cells(I, 9) = "Total Potongan : " & TxtTotPot.Text
        End With

                If CboBahasa.Text = "Indonesia" Then
                    MsgBox "Maaf, Ada kesalahan mengenai pencetakan Data", vbExclamation, "Cetak ke Excel"
                ElseIf CboBahasa.Text = "English" Then
                    MsgBox "Sorry, There was an error about printing the Data", vbExclamation, "Print to Excel"
                End If

                Set ExcelSheet = Nothing
            Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    Me.MousePointer = 99
    Timer2.Enabled = True
    Picture1.Visible = False
End Sub

Private Sub CmdHapus_Click()
    If CboBahasa.Text = "Indonesia" Then
        If (MsgBox("Yakin mau dihapus", vbQuestion + vbOKCancel, "Hapus Data") = vbOK) Then
            ListView1.ListItems.Clear
            CmdTampilkan.Enabled = True
            TxtTotGaji.Text = "0"
            TxtTotPen.Text = "0"
            TxtTotPot.Text = "0"
        End If
    ElseIf CboBahasa.Text = "English" Then
        If (MsgBox("Sure you want to Delete", vbQuestion + vbOKCancel, "Delete Data") = vbOK) Then
            ListView1.ListItems.Clear
            CmdTampilkan.Enabled = True
            TxtTotGaji.Text = "0"
            TxtTotPen.Text = "0"
            TxtTotPot.Text = "0"
        End If
    End If
End Sub

Private Sub CmdInput_Click()
    Call MaxLength_8
    FrmInput.LblRubah.Enabled = False
    FrmInput.LblHapus.Enabled = False
    FrmInput.LblTahun.Caption = Format(Now, "YYYY")
    FrmInput.Show 1
End Sub

Private Sub CmdKeluar_Click()
    If CboBahasa.Text = "Indonesia" Then
        If (MsgBox("Apakah anda yakin ingin Keluar dari Aplikasi ini ?", vbQuestion + vbYesNo, "Keluar") = vbYes) Then
            Call Keluar
        End If
    ElseIf CboBahasa.Text = "English" Then
        If (MsgBox("Do you really wan to Exit this Application ?", vbQuestion + vbYesNo, "Exit") = vbYes) Then
            Call Keluar
        End If
    End If
End Sub

Private Sub CmdTampilkan_Click()
    CmdTampilkan.Enabled = False
    FrmSecurity.Visible = True
    CmdBatal.Enabled = True
    TxtPass.SetFocus
    TxtPass.MaxLength = 20
End Sub

Private Sub Form_Load()
    aHari = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    uHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jum`at", "Sabtu")

    Set ListView1.ColumnHeaderIcons = ImgLst16
    Set ListView1.SmallIcons = ImgLst16

    ListView1.AllowColumnReorder = True
    ListView1.FlatScrollBar = False
'    ListView1.HoverSelection = True
    ListView1.GridLines = True
    ListView1.FullRowSelect = True
    ListView1.View = lvwReport

    CboBahasa.AddItem "English"
    CboBahasa.AddItem "Indonesia"
    CboBahasa.Text = "Indonesia"
    RemoveCancelMenuItem Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CmdKeluar_Click
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'*********
'Sort Asc*
'*********
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True

'***********
'*Sort Desc*
'***********
    If ListView1.Sorted And _
        ColumnHeader.Index - 1 = ListView1.SortKey Then
        ListView1.SortOrder = 1 - ListView1.SortOrder
    Else
        ListView1.SortOrder = lvwAscending
        ListView1.SortKey = ColumnHeader.Index - 1
    End If
        ListView1.Sorted = True
End Sub

Sub Enter()
    If TxtPass.Text = "Bismillah12!" Then
        Picture1.Visible = True
        Timer2.Enabled = True
        Call load_data
        Call Sum_Gaji
        Call CmdBatal_Click
        If ListView1.ListItems.Count = 0 Then
            CmdTampilkan.Enabled = True
            If CboBahasa.Text = "English" Then
                MsgBox "No data yet, please input", vbInformation, "Input Salary"
            ElseIf CboBahasa.Text = "Indonesia" Then
                MsgBox "Belum ada data satupun, silahkan input", vbInformation, "Input Gaji"
            End If
            Call CmdInput_Click
        End If
        Picture1.Visible = False
        Kesempatan = 0
    ElseIf RTrim(LTrim(TxtPass.Text)) = "" Then
        If CboBahasa.Text = "English" Then
            MsgBox "Please retype your Password !", vbExclamation, "Security"
        ElseIf CboBahasa.Text = "Indonesia" Then
            MsgBox "Silahkan Ketik Ulang Kata Sandi Anda !", vbExclamation, "Keamanan"
        End If
    Else
        Kesempatan = Kesempatan + 1
        If CboBahasa.Text = "English" Then
            If Kesempatan = 1 Then
                MsgBox "Incorrect Password, Chance Login (1/3)", vbExclamation, "Security"
            ElseIf Kesempatan = 2 Then
                MsgBox "Incorrect Password, Chance Login (2/3)", vbExclamation, "Security"
            ElseIf Kesempatan = 3 Then
                MsgBox "Incorrect Password, Chance Ends (3/3)", vbExclamation, "Security"
                Call Keluar
            End If
        ElseIf CboBahasa.Text = "Indonesia" Then
            If Kesempatan = 1 Then
                MsgBox "Kata Sandi Salah, Kesempatan Masuk (1/3)", vbExclamation, "Keamanan"
            ElseIf Kesempatan = 2 Then
                MsgBox "Kata Sandi Salah, Kesempatan Masuk (2/3)", vbExclamation, "Keamanan"
            ElseIf Kesempatan = 3 Then
                MsgBox "Kata Sandi Salah, Kesempatan Berakhir (3/3)", vbExclamation, "Keamanan"
                Call Keluar
            End If
        End If
    End If
    TxtPass.Text = ""
End Sub

'Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbRightButton Then
'        Me.PopupMenu mnuWht
'    End If
'End Sub
'
'Private Sub mnuWhat_Click()
'    MsgBox "Run Time Error"
'End Sub
Private Sub Timer3_Timer()
    If FrmSecurity.Visible = True Then
        If GntGmbr = True Then
            StatusBar1.Panels(1).Picture = ImgLst32.ListImages(2).Picture
            GntGmbr = False
        Else
            StatusBar1.Panels(1).Picture = ImgLst32.ListImages(1).Picture
            GntGmbr = True
        End If
    ElseIf ListView1.ListItems.Count > 0 Then
        StatusBar1.Panels(1).Picture = ImgLst32.ListImages(1).Picture
    ElseIf ListView1.ListItems.Count = 0 Then
        StatusBar1.Panels(1).Picture = ImgLst32.ListImages(2).Picture
    End If

    GntIcon = GntIcon + 1
If ListView1.ListItems.Count = 0 Then
    If GntIcon = 1 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(3).Picture
    ElseIf GntIcon = 2 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(4).Picture
    ElseIf GntIcon = 3 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(5).Picture
        GntIcon = 0
    End If
ElseIf ListView1.ListItems.Count > 0 Then
    If GntIcon = 1 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(5).Picture
    ElseIf GntIcon = 2 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(4).Picture
    ElseIf GntIcon = 3 Then
        StatusBar1.Panels(3).Picture = ImgLst32.ListImages(3).Picture
        GntIcon = 0
    End If
End If
'    If RTrim(LTrim(TxtPass.Text)) <> "" Or ChcPass.Value = 1 Then
'        ChcPass.Visible = True
'    ElseIf RTrim(LTrim(TxtPass.Text)) = "" Then
'        ChcPass.Visible = False
'    End If
End Sub

Private Sub Timer1_Timer()
    If CboBahasa.Text = "English" Then
        sHari = aHari(Abs(Weekday(Date) - 1))
        StatusBar1.Panels(4).Text = "" & sHari & ", " _
            & Format(Date, "DD MMMM YYYY") _
            & " - " & Time$
        CboBahasa.ToolTipText = "Laguange"
        StatusBar1.Panels(5).ToolTipText = "Laguange"
    ElseIf CboBahasa.Text = "Indonesia" Then
        sHari = uHari(Abs(Weekday(Date) - 1))
        StatusBar1.Panels(4).Text = "" & sHari & ", " _
            & Format(Date, "DD MMMM YYYY") _
            & " - " & Time
        CboBahasa.ToolTipText = "Bahasa"
        StatusBar1.Panels(5).ToolTipText = "Bahasa"
    End If
        
    If ListView1.ListItems.Count = 0 Then
        If CboBahasa.Text = "Indonesia" Then
            StatusBar1.Panels(2).Text = "Data Kosong"
        ElseIf CboBahasa.Text = "English" Then
            StatusBar1.Panels(2).Text = "Data is Empty"
        End If
        
        CmdInput.Enabled = False
        CmdHapus.Enabled = False
        CmdCari.Enabled = False
        CmdCetak.Enabled = False
        TxtTotGaji.Text = "0"
        TxtTotPot.Text = "0"
        TxtTotPen.Text = "0"
'        If ListView1.ListItems.Count = 0 And CmdBatal.Enabled = False And CmdCari.Enabled = False And CmdHapus.Enabled = False And CmdCetak.Enabled = False Then
'            CmdTampilkan.Enabled = True
'        End If
    ElseIf ListView1.ListItems.Count > 0 Then
        CmdTampilkan.Enabled = False
        If CboBahasa.Text = "Indonesia" Then
            StatusBar1.Panels(2).Text = ListView1.ListItems.Count & " Daftar Data"
            ListView1.ToolTipText = "Klik dua kali untuk detail"
        ElseIf CboBahasa.Text = "English" Then
            StatusBar1.Panels(2).Text = ListView1.ListItems.Count & " List of Data"
            ListView1.ToolTipText = "Double click for details"
        End If
        
        If CmdBatal.Enabled = True And FrmCari.Visible = True Then
            CmdCetak.Enabled = False
            CmdHapus.Enabled = False
            CmdCari.Enabled = False
            CmdInput.Enabled = False
        Else
            CmdCetak.Enabled = True
            CmdHapus.Enabled = True
            CmdCari.Enabled = True
            CmdInput.Enabled = True
        End If
    End If

    StatusBar1.Panels(2).ToolTipText = StatusBar1.Panels(2).Text
    StatusBar1.Panels(4).ToolTipText = StatusBar1.Panels(4).Text
End Sub

Private Sub Timer2_Timer()
    Shape1.Visible = True
    Shape2.Visible = True
    Line1.Visible = False
    Shape2.Width = Shape2.Width + 150
    If Shape2.Width > Shape1.Width Then
        Line1.Visible = True
        Shape1.Visible = False
        Shape2.Visible = False
        Timer2.Enabled = False
        Shape2.Width = 15
    End If
End Sub

Private Sub TxtCari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Cari
    If KeyAscii = vbKeyEscape Then Call CmdBatal_Click
    
    If (KeyAscii = Asc("'")) Then
        TxtCari.Text = Replace(TxtCari.Text, "'", "")
        KeyAscii = 0
        Beep
    End If
End Sub

'Private Sub TxtCari_Change()
'Dim x As Integer
'    ListView1.ListItems.Clear
'        dbConnection
'            rs.Open "Select * From vw_Gaji_Juan Where No = '" & TxtCari.Text & "%'", connect, adOpenDynamic, adLockOptimistic
'                Do Until rs.EOF
'                    Set list = ListView1.ListItems.Add(, , rs(0))
'                        For x = 1 To 8
'                            list.SubItems(x) = rs(x)
'                        Next x
'                    rs.MoveNext
'                Loop
'            Set rs = Nothing
'        connect.Close: Set connect = Nothing
'End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Enter
    If KeyAscii = vbKeyEscape Then Call CmdBatal_Click
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        FrmInput.ImgBulan.Enabled = False
        FrmInput.ImgPerusahaan.Enabled = False
        FrmInput.TxtGaji.Alignment = 1
        FrmInput.TxtPenGapok.Enabled = False
        FrmInput.TxtPenMakan.Enabled = False
        FrmInput.TxtPenTransport.Enabled = False
        FrmInput.TxtPenLembur.Enabled = False
        FrmInput.TxtPenInsHarian.Enabled = False
        FrmInput.TxtPenIns.Enabled = False
        FrmInput.TxtPenJHT.Enabled = False
        FrmInput.TxtPenJKN.Enabled = False
        FrmInput.TxtPenPensiun.Enabled = False
        FrmInput.TxtPenPajak.Enabled = False
        FrmInput.TxtPenLain.Enabled = False
        FrmInput.TxtPenTotal.Alignment = 1
        FrmInput.TxtPotMakan.Enabled = False
        FrmInput.TxtPotJHT.Enabled = False
        FrmInput.TxtPotJKN.Enabled = False
        FrmInput.TxtPotPensiun.Enabled = False
        FrmInput.TxtPotAbsen.Enabled = False
        FrmInput.TxtPotPajak.Enabled = False
        FrmInput.TxtPotLain.Enabled = False
        FrmInput.TxtPotTotal.Alignment = 1
        FrmInput.LblSimpan(0).Enabled = False
        FrmInput.LblForm.Caption = "Detail Salary"
        
        FrmInput.TxtBulan.Text = ListView1.SelectedItem.SubItems(1)
        FrmInput.LblTahun.Caption = ListView1.SelectedItem.SubItems(2)
        FrmInput.TxtPerusahaan.Text = ListView1.SelectedItem.SubItems(3)
        FrmInput.LblJabatan.Caption = "[" & ListView1.SelectedItem.SubItems(3) & "] - [" & ListView1.SelectedItem.SubItems(4) & "]"
        FrmInput.TxtGaji.Text = ListView1.SelectedItem.SubItems(5)
        FrmInput.TxtPenTotal.Text = ListView1.SelectedItem.SubItems(6)
        FrmInput.TxtPotTotal.Text = ListView1.SelectedItem.SubItems(7)
        FrmInput.TxtPenGapok.Text = ListView1.SelectedItem.SubItems(8)
        FrmInput.TxtPenMakan.Text = ListView1.SelectedItem.SubItems(9)
        FrmInput.TxtPenTransport.Text = ListView1.SelectedItem.SubItems(10)
        FrmInput.TxtPenLembur.Text = ListView1.SelectedItem.SubItems(11)
        FrmInput.TxtPenInsHarian.Text = ListView1.SelectedItem.SubItems(12)
        FrmInput.TxtPenIns.Text = ListView1.SelectedItem.SubItems(13)
        FrmInput.TxtPenJHT.Text = ListView1.SelectedItem.SubItems(14)
        FrmInput.TxtPenJKN.Text = ListView1.SelectedItem.SubItems(15)
        FrmInput.TxtPenPensiun.Text = ListView1.SelectedItem.SubItems(16)
        FrmInput.TxtPenPajak.Text = ListView1.SelectedItem.SubItems(17)
        FrmInput.TxtPenLain.Text = ListView1.SelectedItem.SubItems(18)
        FrmInput.TxtPotMakan.Text = ListView1.SelectedItem.SubItems(19)
        FrmInput.TxtPotAbsen.Text = ListView1.SelectedItem.SubItems(20)
        FrmInput.TxtPotJHT.Text = ListView1.SelectedItem.SubItems(21)
        FrmInput.TxtPotJKN.Text = ListView1.SelectedItem.SubItems(22)
        FrmInput.TxtPotPensiun.Text = ListView1.SelectedItem.SubItems(23)
        FrmInput.TxtPotPajak.Text = ListView1.SelectedItem.SubItems(24)
        FrmInput.TxtPotLain.Text = ListView1.SelectedItem.SubItems(25)

        FrmInput.Show 1
    End If
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInput 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8985
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
   Picture         =   "FrmInput.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstPerusahaan 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "FrmInput.frx":10F1D
      Left            =   6390
      List            =   "FrmInput.frx":10F1F
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   4965
      Left            =   7160
      TabIndex        =   17
      Top             =   1700
      Visible         =   0   'False
      Width           =   4640
      Begin VB.TextBox TxtPotTotal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2475
         TabIndex        =   60
         Text            =   "Auto Sum.."
         Top             =   3000
         Width           =   1800
      End
      Begin VB.TextBox TxtPotLain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   59
         Top             =   2640
         Width           =   1800
      End
      Begin VB.TextBox TxtPotPajak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   58
         Top             =   2280
         Width           =   1800
      End
      Begin VB.TextBox TxtPotPensiun 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   57
         Top             =   1920
         Width           =   1800
      End
      Begin VB.TextBox TxtPotJKN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   56
         Top             =   1575
         Width           =   1800
      End
      Begin VB.TextBox TxtPotJHT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   55
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox TxtPotAbsen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   54
         Top             =   840
         Width           =   1800
      End
      Begin VB.TextBox TxtPotMakan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   45
         Top             =   480
         Width           =   1800
      End
      Begin VB.Image Image23 
         Height          =   240
         Left            =   1680
         Picture         =   "FrmInput.frx":10F21
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":11C9B
         Top             =   3000
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   23
         Left            =   2280
         Picture         =   "FrmInput.frx":14EA4
         Top             =   3000
         Width           =   195
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":180A3
         Top             =   2640
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   22
         Left            =   2280
         Picture         =   "FrmInput.frx":1B2AC
         Top             =   2640
         Width           =   195
      End
      Begin VB.Image Image20 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":1E4AB
         Top             =   2280
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   21
         Left            =   2280
         Picture         =   "FrmInput.frx":216B4
         Top             =   2280
         Width           =   195
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":248B3
         Top             =   1920
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   20
         Left            =   2280
         Picture         =   "FrmInput.frx":27ABC
         Top             =   1920
         Width           =   195
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":2ACBB
         Top             =   1575
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   19
         Left            =   2280
         Picture         =   "FrmInput.frx":2DEC4
         Top             =   1575
         Width           =   195
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":310C3
         Top             =   1200
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   18
         Left            =   2280
         Picture         =   "FrmInput.frx":342CC
         Top             =   1200
         Width           =   195
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":374CB
         Top             =   840
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   17
         Left            =   2280
         Picture         =   "FrmInput.frx":3A6D4
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Potongan"
         Height          =   255
         Index           =   24
         Left            =   360
         TabIndex        =   53
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lain - Lain"
         Height          =   255
         Index           =   23
         Left            =   360
         TabIndex        =   52
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pajak PPh 21"
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   51
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jaminan Pensiun"
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   50
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BPJS Kes / JKN"
         Height          =   255
         Index           =   20
         Left            =   360
         TabIndex        =   49
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BPJS TK / JHT"
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   48
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Absen"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   47
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Makan"
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   46
         Top             =   480
         Width           =   1500
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   240
         Y1              =   480
         Y2              =   3240
      End
      Begin VB.Line Line3 
         X1              =   2160
         X2              =   2160
         Y1              =   480
         Y2              =   3240
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   16
         Left            =   2280
         Picture         =   "FrmInput.frx":3D8D3
         Top             =   480
         Width           =   195
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   4275
         Picture         =   "FrmInput.frx":40AD2
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   ".: Tara :."
         Height          =   225
         Index           =   4
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.ListBox LstBulan 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "FrmInput.frx":43CDB
      Left            =   3400
      List            =   "FrmInput.frx":43CDD
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BPJS Kes / "
      Height          =   4965
      Left            =   2520
      TabIndex        =   7
      Top             =   1700
      Visible         =   0   'False
      Width           =   4640
      Begin VB.TextBox TxtPenTotal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2475
         TabIndex        =   44
         Text            =   "Auto Sum.."
         Top             =   4440
         Width           =   1800
      End
      Begin VB.TextBox TxtPenLain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   43
         ToolTipText     =   "PAT / THR, Bonus Perusahaan, Rapel dimasukan Lain-Lain"
         Top             =   4080
         Width           =   1800
      End
      Begin VB.TextBox TxtPenPajak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   42
         Top             =   3720
         Width           =   1800
      End
      Begin VB.TextBox TxtPenPensiun 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   41
         Top             =   3360
         Width           =   1800
      End
      Begin VB.TextBox TxtPenJKN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   40
         Top             =   3000
         Width           =   1800
      End
      Begin VB.TextBox TxtPenJHT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   39
         Top             =   2640
         Width           =   1800
      End
      Begin VB.TextBox TxtPenIns 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   38
         Top             =   2280
         Width           =   1800
      End
      Begin VB.TextBox TxtPenInsHarian 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   37
         Top             =   1920
         Width           =   1800
      End
      Begin VB.TextBox TxtPenLembur 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   36
         ToolTipText     =   "Jika dapat UM Lembur, jumlahkan dengan Lembur"
         Top             =   1575
         Width           =   1800
      End
      Begin VB.TextBox TxtPenTransport 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   35
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox TxtPenMakan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2475
         TabIndex        =   34
         Top             =   840
         Width           =   1800
      End
      Begin VB.TextBox TxtPenGapok 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2480
         TabIndex        =   21
         Top             =   480
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   24
         Left            =   1680
         Picture         =   "FrmInput.frx":43CDF
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":44A59
         Top             =   4440
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   15
         Left            =   2280
         Picture         =   "FrmInput.frx":47C62
         Top             =   4440
         Width           =   195
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":4AE61
         ToolTipText     =   "PAT / THR, Bonus Perusahaan, Rapel dimasukan Lain-Lain"
         Top             =   4080
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   14
         Left            =   2280
         Picture         =   "FrmInput.frx":4E06A
         ToolTipText     =   "PAT / THR, Bonus Perusahaan, Rapel dimasukan Lain-Lain"
         Top             =   4080
         Width           =   195
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":51269
         Top             =   3720
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   13
         Left            =   2280
         Picture         =   "FrmInput.frx":54472
         Top             =   3720
         Width           =   195
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":57671
         Top             =   3360
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   12
         Left            =   2280
         Picture         =   "FrmInput.frx":5A87A
         Top             =   3360
         Width           =   195
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":5DA79
         Top             =   3000
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   11
         Left            =   2280
         Picture         =   "FrmInput.frx":60C82
         Top             =   3000
         Width           =   195
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":63E81
         Top             =   2640
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   10
         Left            =   2280
         Picture         =   "FrmInput.frx":6708A
         Top             =   2640
         Width           =   195
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":6A289
         Top             =   2280
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   9
         Left            =   2280
         Picture         =   "FrmInput.frx":6D492
         Top             =   2280
         Width           =   195
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":70691
         Top             =   1920
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   8
         Left            =   2280
         Picture         =   "FrmInput.frx":7389A
         Top             =   1920
         Width           =   195
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":76A99
         ToolTipText     =   "Jika dapat UM Lembur, jumlahkan dengan Lembur"
         Top             =   1575
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   7
         Left            =   2280
         Picture         =   "FrmInput.frx":79CA2
         ToolTipText     =   "Jika dapat UM Lembur, jumlahkan dengan Lembur"
         Top             =   1575
         Width           =   195
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":7CEA1
         Top             =   1200
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   6
         Left            =   2280
         Picture         =   "FrmInput.frx":800AA
         Top             =   1200
         Width           =   195
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":832A9
         Top             =   840
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   5
         Left            =   2280
         Picture         =   "FrmInput.frx":864B2
         Top             =   840
         Width           =   195
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   2160
         Y1              =   480
         Y2              =   4680
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   480
         Y2              =   4680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pendapatan"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   33
         Top             =   4440
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lain - Lain"
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   32
         ToolTipText     =   "PAT / THR, Bonus Perusahaan, Rapel dimasukan Lain-Lain"
         Top             =   4080
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pajak PPh 21"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   31
         Top             =   3720
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jaminan Pensiun"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   30
         Top             =   3360
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BPJS Kes / JKN"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   29
         Top             =   3000
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BPJS TK / JHT"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   28
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insentif"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insentif Harian"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   26
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Lembur"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   25
         ToolTipText     =   "Jika dapat UM Lembur, jumlahkan dengan Lembur"
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Transport"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Makan"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji Pokok"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   1140
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   2280
         Picture         =   "FrmInput.frx":896B1
         Top             =   480
         Width           =   195
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   4280
         Picture         =   "FrmInput.frx":8C8B0
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   ".: Bruto :."
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1005
      Left            =   2500
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   9280
      Begin VB.TextBox TxtGaji 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   7515
         TabIndex        =   15
         Text            =   "Auto Sum.."
         Top             =   480
         Width           =   1530
      End
      Begin VB.TextBox TxtPerusahaan 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   4040
         TabIndex        =   11
         Top             =   480
         Width           =   1530
      End
      Begin VB.TextBox TxtBulan 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1065
         TabIndex        =   9
         Top             =   480
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   9000
         Picture         =   "FrmInput.frx":8FAB9
         Top             =   480
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   3
         Left            =   7320
         Picture         =   "FrmInput.frx":92CC2
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji Dibayar"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   14
         Top             =   480
         Width           =   1260
      End
      Begin VB.Image ImgPerusahaan 
         Height          =   255
         Left            =   5565
         Picture         =   "FrmInput.frx":95EC1
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   2
         Left            =   3840
         Picture         =   "FrmInput.frx":991D4
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   660
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   1
         Left            =   870
         Picture         =   "FrmInput.frx":9C3D3
         Top             =   480
         Width           =   195
      End
      Begin VB.Image ImgBulan 
         Height          =   255
         Left            =   2235
         Picture         =   "FrmInput.frx":9F5D2
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   ".: Netto :."
         Height          =   225
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9345
      End
   End
   Begin VB.Timer Timer5 
      Left            =   8640
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Left            =   8160
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Left            =   7680
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7770
      Left            =   165
      TabIndex        =   0
      Top             =   645
      Width           =   2130
      Begin VB.Label LblBatal 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "   Bersihkan"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1470
         Width           =   1875
      End
      Begin VB.Image ImgBatal 
         Height          =   375
         Left            =   45
         Picture         =   "FrmInput.frx":A28E5
         Top             =   1400
         Width           =   2055
      End
      Begin VB.Label LblSimpan 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "   Simpan"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   280
         Width           =   1875
      End
      Begin VB.Image ImgSimpan 
         Height          =   375
         Index           =   0
         Left            =   45
         Picture         =   "FrmInput.frx":A5A5C
         Top             =   200
         Width           =   2055
      End
      Begin VB.Label LblRubah 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "   Rubah Data"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   690
         Width           =   1875
      End
      Begin VB.Image ImgRubah 
         Height          =   375
         Left            =   45
         Picture         =   "FrmInput.frx":A8BD3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label LblHapus 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "   Hapus Data"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Image ImgHapus 
         Height          =   375
         Left            =   45
         Picture         =   "FrmInput.frx":ABD4A
         Tag             =   "1"
         Top             =   1000
         Width           =   2055
      End
      Begin VB.Image ImgBackMenuPane 
         Height          =   7785
         Left            =   0
         Picture         =   "FrmInput.frx":AEEC1
         Top             =   0
         Width           =   2145
      End
   End
   Begin VB.Timer Timer2 
      Left            =   8760
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Left            =   8280
      Top             =   120
   End
   Begin MSComctlLib.ListView LvwKry 
      Height          =   1750
      Left            =   2520
      TabIndex        =   19
      ToolTipText     =   "Klik dua kali untuk detail"
      Top             =   6675
      Visible         =   0   'False
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   3096
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16382457
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
      Picture         =   "FrmInput.frx":B3D5D
   End
   Begin VB.Label LblJabatan 
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   64
      ToolTipText     =   "Jabatan"
      Top             =   8745
      Width           =   10335
   End
   Begin VB.Label LblTahun 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YYYY"
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
      Left            =   10680
      TabIndex        =   61
      ToolTipText     =   "Tahun Periode Gaji"
      Top             =   8740
      Width           =   1215
   End
   Begin VB.Label LblMax 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   11315
      TabIndex        =   16
      ToolTipText     =   "None"
      Top             =   105
      Width           =   165
   End
   Begin VB.Label LblExit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   11660
      TabIndex        =   4
      ToolTipText     =   "Close"
      Top             =   105
      Width           =   165
   End
   Begin VB.Label LblMini 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   160
      Left            =   10940
      TabIndex        =   3
      ToolTipText     =   "Minimize"
      Top             =   100
      Width           =   160
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   10815
      Picture         =   "FrmInput.frx":B7EC1
      Top             =   15
      Width           =   1125
   End
   Begin VB.Label LblForm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Input Salary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   60
      Width           =   9570
   End
End
Attribute VB_Name = "FrmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim posTengah As Integer
Dim GntBln As Boolean
Dim GntPrs As Boolean
Dim JmlPersen As Integer

Private Sub ImgPerusahaan_Click()
    LstBulan.Visible = False
    If GntPrs = True Then
        LstPerusahaan.Visible = False
        GntPrs = False
    Else
        LstPerusahaan.Visible = True
        GntPrs = True
    End If
End Sub

Private Sub ImgBulan_Click()
    LstPerusahaan.Visible = False
    If GntBln = True Then
        LstBulan.Visible = False
        GntBln = False
    Else
        LstBulan.Visible = True
        GntBln = True
    End If
End Sub

Private Sub Form_Load()
'    Dim MenuSamping(10) As String

    'mengisi label menu
'    For i = 0 To 10
'        MenuSamping(i) = "   Menu Ke " & i + 1
'    Next i
    
    Me.Top = Me.Height * -1
    Me.Left = (Screen.Width - Me.Width) / 2
    posTengah = (Screen.Height - Me.Height) / 2
    Me.Timer1.Interval = 50
    Me.Frame1.Left = Me.Frame1.Width * -1

'   For i = 0 To 10 ' jumlah menu
'       Load Me.ImgSimpan(100 + i)
'       Me.ImgSimpan(100 + i).Picture = LoadPicture(App.Path & "\Images\MnButtonHover.jpg")
        If I > 0 Then
            Load Me.LblSimpan(I)
            Me.LblSimpan(I).Visible = True
            Load Me.ImgSimpan(I)
            Me.ImgSimpan(I).Visible = True
        End If
        Me.LblSimpan(I).Left = Me.LblSimpan(0).Width * -1
        Me.LblSimpan(I).Top = Me.LblSimpan(0).Top
'       Me.LblSimpan(i).Caption = MenuSamping(i)
        Me.ImgSimpan(I).Left = Me.ImgSimpan(I).Width * -1
        Me.ImgSimpan(I).Top = Me.LblSimpan(I).Top - 70
'   Next i

    Me.LblRubah.Left = Me.LblSimpan(0).Left
    Me.ImgRubah.Left = Me.ImgSimpan(0).Left
    Me.LblHapus.Left = Me.LblSimpan(0).Left
    Me.ImgHapus.Left = Me.ImgSimpan(0).Left
    Me.LblBatal.Left = Me.LblSimpan(0).Left
    Me.ImgBatal.Left = Me.ImgSimpan(0).Left

'    Me.ImgBackMenuPane.ZOrder 1
'    Me.Frame2.Left = Me.Frame2.Width * -1

    LvwKry.ColumnHeaders.Add , , "No", 500
    LvwKry.ColumnHeaders.Add , , "NIK", 1300
    LvwKry.ColumnHeaders.Add , , "Perusahaan", 2500
    LvwKry.ColumnHeaders.Add , , "Bagian", 1500
    LvwKry.ColumnHeaders.Add , , "Jabatan", 1500
    LvwKry.ColumnHeaders.Add , , "Tanggal Masuk", 1700
    LvwKry.ColumnHeaders.Add , , "Tanggal Keluar", 1700
    LvwKry.ColumnHeaders.Add , , "Status", 1300
    LvwKry.ColumnHeaders.Add , , "BPJS TK (JHT)", 900
    LvwKry.ColumnHeaders.Add , , "BPJS Kes (JKN)", 900
    LvwKry.ColumnHeaders.Add , , "NPWP", 900
    LvwKry.ColumnHeaders.Add , , "Alamat", 4000
    
    Call ComboBox
    Call Pekerjaan
    
    LvwKry.AllowColumnReorder = True
    LvwKry.FlatScrollBar = False
    LvwKry.HoverSelection = True
    LvwKry.GridLines = True
    LvwKry.FullRowSelect = True
    LvwKry.View = lvwReport
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, _
        WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
End Sub

Private Sub ImgBackMenuPane_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Me.ImgBatal.Tag = 2 Then
'        Me.ImgBatal.Picture = LoadPicture(App.Path & "\Images\MnButton.jpg")
'        Me.ImgBatal.Tag = 1
'    End If

    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub LblBatal_Click()
    If LblBatal.Caption = "   Bersihkan" Then
        Call Bersihkan
        Call MaxLength_8
    ElseIf LblBatal.Caption = "   Batal" Then
        Call MaxLength_13
        Call TampilRP
        Call EnabledFalse
        LblSimpan(0).Caption = "   Simpan"
    End If
End Sub

Private Sub LblHapus_Click()
    LblForm.Caption = "Delete Salary"
    If (MsgBox("Yakin data Gaji " & TxtBulan.Text & " " & LblTahun.Caption & " dari " & TxtPerusahaan.Text & " ingin dihapus?", vbQuestion + vbYesNo, "Delete Salary") = vbYes) Then
        Call HapusData
        Call EnabledTrue
        Call Bersihkan
        MsgBox "Data berhasil dihapus..", vbInformation, "Delete Salary"
        Call LblExit_Click
    ElseIf vbNo Then
        LblForm.Caption = "Detail Salary"
    End If
End Sub

Private Sub LblRubah_Click()
    Call MaxLength_8
    Call TampilNotRP
    Call EnabledTrue
    LblForm.Caption = "Edit Salary"
    LblSimpan(0).Caption = "   Simpan Peruba.."
    LblBatal.Caption = "   Batal"
End Sub

Sub EnabledTrue()
    ImgBulan.Enabled = True
    ImgPerusahaan.Enabled = True
    TxtGaji.Alignment = 0
    TxtGaji.Text = "Auto Sum.."
    TxtPenGapok.Enabled = True
    TxtPenMakan.Enabled = True
    TxtPenTransport.Enabled = True
    TxtPenLembur.Enabled = True
    TxtPenInsHarian.Enabled = True
    TxtPenIns.Enabled = True
    TxtPenJHT.Enabled = True
    TxtPenJKN.Enabled = True
    TxtPenPensiun.Enabled = True
    TxtPenPajak.Enabled = True
    TxtPenLain.Enabled = True
    TxtPenTotal.Alignment = 0
    TxtPenTotal.Text = "Auto Sum.."
    TxtPotMakan.Enabled = True
    TxtPotJHT.Enabled = True
    TxtPotJKN.Enabled = True
    TxtPotPensiun.Enabled = True
    TxtPotAbsen.Enabled = True
    TxtPotPajak.Enabled = True
    TxtPotLain.Enabled = True
    TxtPotTotal.Alignment = 0
    TxtPotTotal.Text = "Auto Sum.."
    LblHapus.Enabled = False
    LblRubah.Enabled = False
    LblSimpan(0).Enabled = True
End Sub

Sub EnabledFalse()
    ImgBulan.Enabled = False
    ImgPerusahaan.Enabled = False
    TxtGaji.Alignment = 1
    TxtPenGapok.Enabled = False
    TxtPenMakan.Enabled = False
    TxtPenTransport.Enabled = False
    TxtPenLembur.Enabled = False
    TxtPenInsHarian.Enabled = False
    TxtPenIns.Enabled = False
    TxtPenJHT.Enabled = False
    TxtPenJKN.Enabled = False
    TxtPenPensiun.Enabled = False
    TxtPenPajak.Enabled = False
    TxtPenLain.Enabled = False
    TxtPenTotal.Alignment = 1
    TxtPotMakan.Enabled = False
    TxtPotJHT.Enabled = False
    TxtPotJKN.Enabled = False
    TxtPotPensiun.Enabled = False
    TxtPotAbsen.Enabled = False
    TxtPotPajak.Enabled = False
    TxtPotLain.Enabled = False
    TxtPotTotal.Alignment = 1
    LblSimpan(0).Enabled = False
    LblRubah.Enabled = True
    LblHapus.Enabled = True
    LblBatal.Caption = "   Bersihkan"
    LblForm.Caption = "Detail Salary"
End Sub

Sub Bersihkan()
    ImgBulan.Enabled = True
    ImgPerusahaan.Enabled = True
    TxtGaji.Alignment = 0
    TxtPenGapok.Enabled = True
    TxtPenMakan.Enabled = True
    TxtPenTransport.Enabled = True
    TxtPenLembur.Enabled = True
    TxtPenInsHarian.Enabled = True
    TxtPenIns.Enabled = True
    TxtPenJHT.Enabled = True
    TxtPenJKN.Enabled = True
    TxtPenPensiun.Enabled = True
    TxtPenPajak.Enabled = True
    TxtPenLain.Enabled = True
    TxtPenTotal.Alignment = 0
    TxtPotMakan.Enabled = True
    TxtPotJHT.Enabled = True
    TxtPotJKN.Enabled = True
    TxtPotPensiun.Enabled = True
    TxtPotAbsen.Enabled = True
    TxtPotPajak.Enabled = True
    TxtPotLain.Enabled = True
    TxtPotTotal.Alignment = 0
    LblSimpan(0).Enabled = True
    LblRubah.Enabled = False
    LblHapus.Enabled = False
    LblBatal.Caption = "   Bersihkan"
    LblTahun.Caption = Format(Now, "YYYY")
    LblForm.Caption = "Input Salary"
    LstBulan.Visible = False
    LstPerusahaan.Visible = False
    GntPrs = False
    GntBln = False

    TxtBulan.Text = ""
    TxtPerusahaan.Text = ""
    TxtGaji.Text = "Auto Sum.."
    TxtGaji.Alignment = 0
    TxtPenGapok.Text = ""
    TxtPenMakan.Text = ""
    TxtPenTransport.Text = ""
    TxtPenLembur.Text = ""
    TxtPenInsHarian.Text = ""
    TxtPenIns.Text = ""
    TxtPenJHT.Text = ""
    TxtPenJKN.Text = ""
    TxtPenPensiun.Text = ""
    TxtPenPajak.Text = ""
    TxtPenLain.Text = ""
    TxtPenTotal.Text = "Auto Sum.."
    TxtPenTotal.Alignment = 0
    TxtPotMakan.Text = ""
    TxtPotJHT.Text = ""
    TxtPotJKN.Text = ""
    TxtPotPensiun.Text = ""
    TxtPotAbsen.Text = ""
    TxtPotPajak.Text = ""
    TxtPotLain.Text = ""
    TxtPotTotal.Text = "Auto Sum.."
    TxtPotTotal.Alignment = 0
    LblJabatan.Caption = ""
    TxtPenGapok.SetFocus
End Sub

Private Sub LblSimpan_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    For i = 0 To 10
'        Me.ImgSimpan(100 + i).Visible = False
'        Me.ImgSimpan(i).Visible = True
'    Next i
'    Me.ImgSimpan(100 + Index).Visible = True
'    Me.ImgSimpan(Index).Visible = False
'    Me.LblSimpan(Index).ZOrder 0
    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(7).Picture
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub LblRubah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(7).Picture
    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub LblHapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(7).Picture
    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub LblBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Me.ImgBatal.Tag = 1 Then
'        Me.ImgBatal.Picture = LoadPicture(App.Path & "\Images\MnButtonHover.jpg")
'        Me.ImgBatal.Tag = 2
'    End If
    ImgBatal.Picture = FrmGaji.ImgLst32.ListImages(7).Picture
    ImgSimpan(Index).Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgRubah.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
    ImgHapus.Picture = FrmGaji.ImgLst32.ListImages(6).Picture
End Sub

Private Sub LblForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, _
        WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
End Sub

Private Sub LblExit_Click()
    If LblForm.Caption = "Edit Salary" Then
        MsgBox "Simpan / Batalkan terlebih dahulu data yg ingin dirubah !", vbExclamation, LblForm.Caption
    Else
        posTengah = Me.Top
        Me.Top = Me.Top + 10
        Me.Timer2.Interval = 50
        Call load_data
        Call Sum_Gaji
        
        If FrmGaji.ListView1.ListItems.Count = 0 Then
            FrmGaji.CmdTampilkan.Enabled = True
        End If
    End If
End Sub

Private Sub LblMax_Click()
    Beep
End Sub

Private Sub LblMini_Click()
    Me.WindowState = vbMinimized
    FrmGaji.WindowState = vbMinimized
End Sub

Private Sub LblSimpan_Click(Index As Integer)
    If LblSimpan(0).Caption = "   Simpan" Then
        If TxtBulan.Text = "" Or TxtPerusahaan.Text = "" Or TxtGaji.Text = "Auto Sum.." Or TxtPenTotal.Text = "Auto Sum.." Or TxtPotTotal.Text = "Auto Sum.." Then
            MsgBox "Lengkapi data terlebih dahulu !", vbExclamation, "Input Data"
        Else
            Call SimpanData
            Call Bersihkan
            MsgBox "Data berhasil disimpan..", vbInformation, "Input Data"
            Call LblExit_Click
        End If
    ElseIf LblSimpan(0).Caption = "   Simpan Peruba.." Then
        Call RubahData
        Call Bersihkan
        MsgBox "Data berhasil dirubah..", vbInformation, "Input Data"
        Call LblExit_Click
    End If
End Sub

Private Sub LstPerusahaan_Click()
    TxtPerusahaan.Text = LstPerusahaan.Text
    LstPerusahaan.Visible = False
    GntPrs = False
    Call StatusKry
End Sub

Private Sub LstBulan_Click()
    TxtBulan.Text = LstBulan.Text
    LstBulan.Visible = False
    GntBln = False
End Sub

Private Sub LvwKry_DblClick()
'    Call M_Pekerjaan
'    Timer5.Enabled = False
'    Frame2.Visible = False
'    Frame3.Visible = False
'    Frame4.Visible = False
'    LvwKry.Visible = False
    
    Select Case Index
    Case 0
    
    Call M_Pekerjaan
    Timer5.Enabled = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    LvwKry.Visible = False
        With FrmPekerjaan
            .Top = Me.Top + Me.Frame1.Top
            .Left = Me.Left + Me.Frame1.Left + Me.Frame1.Width + 300
            .Show 1
'            .Visible = True
        End With
    End Select
End Sub

Private Sub Timer1_Timer()
    If Me.Top < posTengah Then
        Me.Top = Me.Top + ((posTengah - Me.Top) / 2)
    Else
        Me.Top = posTengah
        Me.Timer1.Interval = 0
        Me.Timer3.Interval = 50
    End If
End Sub

Private Sub Timer2_Timer()
    If Me.Top < Screen.Height Then
        Me.Top = Me.Top + ((Me.Top - posTengah) * 2)
    Else
        Unload Me
    End If
End Sub

Private Sub Timer3_Timer()
    If Me.Frame1.Left < 120 Then
        Me.Frame1.Left = Me.Frame1.Left + ((120 - Me.Frame1.Left) / 2)
        If Me.Frame1.Left > 50 Then Me.Timer4.Interval = 50
    Else
        Me.Frame1.Left = 130
        Me.Refresh
        DoEvents
        Me.Frame1.Left = 120
        Me.Refresh
        Me.Timer3.Interval = 0
    End If
End Sub

Private Sub Timer4_Timer()
'   For i = 0 To 10
        If Me.LblSimpan(I).Left < 120 Then
            Me.LblSimpan(I).Left = Me.LblSimpan(I).Left + ((120 - Me.LblSimpan(I).Left) / 2)
            Me.ImgSimpan(I).Left = Me.LblSimpan(I).Left - 70
'           If Me.LblSimpan(7).Left > 90 Then Me.Timer5.Interval = 50
        Else
            Me.LblSimpan(I).Left = 130
            Me.Refresh
            DoEvents
            Me.LblSimpan(I).Left = 120
            Me.Refresh
            Me.Timer4.Interval = 0
        End If
'   Next i
    
    Me.LblRubah.Left = Me.LblSimpan(0).Left
    Me.ImgRubah.Left = Me.ImgSimpan(0).Left
    Me.LblHapus.Left = Me.LblSimpan(0).Left
    Me.ImgHapus.Left = Me.ImgSimpan(0).Left
    Me.LblBatal.Left = Me.LblSimpan(0).Left
    Me.ImgBatal.Left = Me.ImgSimpan(0).Left
    Me.Timer5.Interval = 500
End Sub

Private Sub Timer5_Timer()
'    For i = 1 To 10
'        If Me.LblSimpan(i).Top < 120 + 360 * i Then
'            Me.LblSimpan(i).Top = Me.LblSimpan(i).Top + (((120 + 360 * i) - Me.LblSimpan(i).Top) / 2)
'            Me.ImgSimpan(i).Top = Me.LblSimpan(i).Top - 70
'        Else
'            Me.LblSimpan(i).Top = 120 + 360 * i
'            Me.ImgSimpan(100 + i).Top = Me.ImgSimpan(i).Top
'            Me.ImgSimpan(100 + i).Left = Me.ImgSimpan(i).Left
'            If i = 7 Then Me.Timer5.Interval = 0
'        End If
'    Next i
'
    Frame2.Visible = True
    Frame3.Visible = True
    Frame4.Visible = True
    LvwKry.Visible = True

    If LblSimpan(0).Enabled = True Then
        If TxtPenGapok.Text = "" And TxtPenMakan.Text = "" And TxtPenTransport.Text = "" And TxtPenLembur.Text = "" And TxtPenInsHarian.Text = "" And TxtPenIns.Text = "" And TxtPenJHT.Text = "" And TxtPenJKN.Text = "" And TxtPenPensiun.Text = "" And TxtPenPajak.Text = "" And TxtPenLain.Text = "" Then
            TxtPenTotal.Text = "Auto Sum.."
            TxtPenTotal.Alignment = 0
        Else
            TxtPenTotal.Text = Val(TxtPenGapok.Text) + Val(TxtPenMakan.Text) + Val(TxtPenTransport.Text) + Val(TxtPenLembur.Text) + Val(TxtPenInsHarian.Text) + Val(TxtPenIns.Text) + Val(TxtPenJHT.Text) + Val(TxtPenJKN.Text) + Val(TxtPenPensiun.Text) + Val(TxtPenPajak.Text) + Val(TxtPenLain.Text)
            TxtPenTotal.Alignment = 1
        End If

        If TxtPotMakan.Text = "" And TxtPotJHT.Text = "" And TxtPotJKN.Text = "" And TxtPotPensiun.Text = "" And TxtPotAbsen.Text = "" And TxtPotPajak.Text = "" And TxtPotLain.Text = "" Then
            TxtPotTotal.Text = "Auto Sum.."
            TxtPotTotal.Alignment = 0
            TxtGaji.Alignment = 0
        Else
            TxtPotTotal.Text = Val(TxtPotMakan.Text) + Val(TxtPotJHT.Text) + Val(TxtPotJKN.Text) + Val(TxtPotPensiun.Text) + Val(TxtPotAbsen.Text) + Val(TxtPotPajak.Text) + Val(TxtPotLain.Text)
            TxtGaji.Text = Val(TxtPenTotal.Text) - Val(TxtPotTotal.Text)
            TxtPotTotal.Alignment = 1
            TxtGaji.Alignment = 1
        End If
    End If
    
    If ImgBulan.Enabled = False Or ImgPerusahaan.Enabled = False Then
        LstBulan.Visible = False
        LstPerusahaan.Visible = False
    End If
End Sub

'------------ Numerik Text ------------
Private Sub TxtPenGapok_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenMakan_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenTransport_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenLembur_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenInsHarian_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenIns_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenJHT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenJKN_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenPensiun_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenPajak_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPenLain_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotMakan_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotJHT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotJKN_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotPensiun_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotAbsen_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotPajak_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPotLain_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        MsgBox "Harap gunakan angka !", vbCritical, LblForm.Caption
        KeyAscii = 0
    End If
End Sub

'------------ Jika kosong maka Text Box jadi 0 ------------
Private Sub TxtPenGapok_LostFocus()
    If LTrim(RTrim(TxtPenGapok)) = "" Then
        TxtPenGapok.Text = "0"
    End If
End Sub

Private Sub TxtPenMakan_LostFocus()
    If LTrim(RTrim(TxtPenMakan.Text)) = "" Then
        TxtPenMakan.Text = "0"
    End If
End Sub

Private Sub TxtPenTransport_LostFocus()
    If LTrim(RTrim(TxtPenTransport.Text)) = "" Then
        TxtPenTransport.Text = "0"
    End If
End Sub

Private Sub TxtPenLembur_LostFocus()
    If LTrim(RTrim(TxtPenLembur)) = "" Then
        TxtPenLembur.Text = "0"
    End If
End Sub

Private Sub TxtPenInsHarian_LostFocus()
    If LTrim(RTrim(TxtPenInsHarian.Text)) = "" Then
        TxtPenInsHarian.Text = "0"
    End If
End Sub

Private Sub TxtPenIns_LostFocus()
    If LTrim(RTrim(TxtPenIns.Text)) = "" Then
        TxtPenIns.Text = "0"
    End If
End Sub

Private Sub TxtPenJHT_LostFocus()
    If LTrim(RTrim(TxtPenJHT.Text)) = "" Then
        TxtPenJHT.Text = "0"
    End If
End Sub

Private Sub TxtPenJKN_LostFocus()
    If LTrim(RTrim(TxtPenJKN.Text)) = "" Then
        TxtPenJKN.Text = "0"
    End If
End Sub

Private Sub TxtPenPensiun_LostFocus()
    If LTrim(RTrim(TxtPenPensiun.Text)) = "" Then
        TxtPenPensiun.Text = "0"
    End If
End Sub

Private Sub TxtPenPajak_LostFocus()
    If LTrim(RTrim(TxtPenPajak.Text)) = "" Then
        TxtPenPajak.Text = "0"
    End If
End Sub

Private Sub TxtPenLain_LostFocus()
    If LTrim(RTrim(TxtPenLain.Text)) = "" Then
        TxtPenLain.Text = "0"
    End If
End Sub

Private Sub TxtPotMakan_LostFocus()
    If LTrim(RTrim(TxtPotMakan.Text)) = "" Then
        TxtPotMakan.Text = "0"
    End If
End Sub

Private Sub TxtPotJHT_LostFocus()
    If LTrim(RTrim(TxtPotJHT.Text)) = "" Then
        TxtPotJHT.Text = "0"
    End If
End Sub

Private Sub TxtPotJKN_LostFocus()
    If LTrim(RTrim(TxtPotJKN.Text)) = "" Then
        TxtPotJKN.Text = "0"
    End If
End Sub

Private Sub TxtPotPensiun_LostFocus()
    If LTrim(RTrim(TxtPotPensiun.Text)) = "" Then
        TxtPotPensiun.Text = "0"
    End If
End Sub

Private Sub TxtPotAbsen_LostFocus()
    If LTrim(RTrim(TxtPotAbsen.Text)) = "" Then
        TxtPotAbsen.Text = "0"
    End If
End Sub

Private Sub TxtPotPajak_LostFocus()
    If LTrim(RTrim(TxtPotPajak.Text)) = "" Then
        TxtPotPajak.Text = "0"
    End If
End Sub

Private Sub TxtPotLain_LostFocus()
    If LTrim(RTrim(TxtPotLain.Text)) = "" Then
        TxtPotLain.Text = "0"
    End If
End Sub

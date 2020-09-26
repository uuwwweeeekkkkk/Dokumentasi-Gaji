Attribute VB_Name = "MdlDB"

Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim list As ListItem
Dim InsSQL As String
Dim DelSQL As String
Dim UpdSQL As String
Dim KodeBulan As String
   
Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub RemoveCancelMenuItem(frm As Form)
Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = -1
End Sub

Sub dbConnection()
    If connect.State = 1 Then
        connect.Close
    End If
    
    If rs.State = 1 Then
        rs.Close
    End If
    
    connect.Open "Provider=SQLOLEDB.1;Persist Security Info=False;server=(local);database=Dev_Dokumentasi;uid=sa;pwd=sa"
End Sub

Sub load_data()
    FrmGaji.ListView1.ListItems.Clear
    dbConnection
        rs.Open "Select * From vw_Gaji_Juan", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            Set list = FrmGaji.ListView1.ListItems.Add(, , FrmGaji.ListView1.ListItems.Count + 1, , 2)
                list.SubItems(1) = rs!Bulan
                list.SubItems(2) = rs!Tahun
                list.SubItems(3) = rs!Perusahaan
                list.SubItems(4) = rs!Jabatan
                list.SubItems(5) = Format(Val(rs!GajiYangDibayar), "Rp ###,##,0")
                list.SubItems(6) = Format(Val(rs!PEN_Total_Pendapatan), "Rp ###,##,0")
                list.SubItems(7) = Format(Val(rs!POT_Total_Potongan), "Rp ###,##,0")
                list.SubItems(8) = Format(Val(rs!PEN_Gaji_Pokok_Bulanan), "Rp ###,##,0")
                list.SubItems(9) = Format(Val(rs!PEN_Uang_Makan), "Rp ###,##,0")
                list.SubItems(10) = Format(Val(rs!PEN_Uang_Transportasi), "Rp ###,##,0")
                list.SubItems(11) = Format(Val(rs!PEN_Uang_Lembur), "Rp ###,##,0")
                list.SubItems(12) = Format(Val(rs!PEN_Insentif_Harian), "Rp ###,##,0")
                list.SubItems(13) = Format(Val(rs!PEN_Insentif), "Rp ###,##,0")
                list.SubItems(14) = Format(Val(rs!PEN_JHT), "Rp ###,##,0")
                list.SubItems(15) = Format(Val(rs!PEN_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(16) = Format(Val(rs!PEN_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(17) = Format(Val(rs!PEN_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(18) = Format(Val(rs!PEN_Lain), "Rp ###,##,0")
                list.SubItems(19) = Format(Val(rs!POT_Uang_Makan), "Rp ###,##,0")
                list.SubItems(20) = Format(Val(rs!POT_Absen), "Rp ###,##,0")
                list.SubItems(21) = Format(Val(rs!POT_JHT), "Rp ###,##,0")
                list.SubItems(22) = Format(Val(rs!POT_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(23) = Format(Val(rs!POT_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(24) = Format(Val(rs!POT_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(25) = Format(Val(rs!POT_Lain), "Rp ###,##,0")
                list.SubItems(26) = Format(rs!Tanggal, "YYYY-MM-DD HH:MM")
            rs.MoveNext
        Loop
End Sub

Sub Sum_Gaji()
    FrmGaji.TxtTotGaji.Text = "0"
    FrmGaji.TxtTotPen.Text = "0"
    FrmGaji.TxtTotPot.Text = "0"
    dbConnection

    If FrmGaji.CboBahasa.Text = "English" Then
    rs.Open "Select IsNull(Sum(GajiYangDibayar), '0') As [Total Gaji Dibayar]," & vbfrlf & _
            "IsNull(Sum(PEN_Total_Pendapatan), '0') As [Total Pendapatan]," & vbCrLf & _
            "IsNull(Sum(POT_Total_Potongan), '0') As [Total Potongan]" & vbCrLf & _
            "From vw_Gaji_Juan_Eng" & vbCrLf & _
            "Where " & FrmGaji.CboCari.Text & " Like '%" & FrmGaji.TxtCari.Text & "%'", connect, adOpenDynamic, adLockOptimistic
    ElseIf FrmGaji.CboBahasa.Text = "Indonesia" Then
    rs.Open "Select IsNull(Sum(GajiYangDibayar), '0') As [Total Gaji Dibayar]," & vbfrlf & _
            "IsNull(Sum(PEN_Total_Pendapatan), '0') As [Total Pendapatan]," & vbCrLf & _
            "IsNull(Sum(POT_Total_Potongan), '0') As [Total Potongan]" & vbCrLf & _
            "From vw_Gaji_Juan" & vbCrLf & _
            "Where " & FrmGaji.CboCari.Text & " Like '%" & FrmGaji.TxtCari.Text & "%'", connect, adOpenDynamic, adLockOptimistic
    End If
        Do Until rs.EOF
            FrmGaji.TxtTotGaji.Text = Format(Val(rs![Total Gaji Dibayar]), "Rp ###,##,0")
            FrmGaji.TxtTotPen.Text = Format(Val(rs![Total Pendapatan]), "Rp ###,##,0")
            FrmGaji.TxtTotPot.Text = Format(Val(rs![Total Potongan]), "Rp ###,##,0")
        rs.MoveNext
        Loop
End Sub

Sub Cari()
    If FrmGaji.CboBahasa.Text = "Indonesia" Then
        Call Cari_Ind
    ElseIf FrmGaji.CboBahasa.Text = "English" Then
        Call Cari_Eng
    End If

    Call Sum_Gaji
    Call Bersih_Cari
End Sub

Sub Cari_Ind()
    FrmGaji.ListView1.ListItems.Clear
    dbConnection
        rs.Open "Select * From vw_Gaji_Juan" & vbCrLf & _
                "Where " & FrmGaji.CboCari.Text & " Like '%" & FrmGaji.TxtCari.Text & "%'", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            Set list = FrmGaji.ListView1.ListItems.Add(, , FrmGaji.ListView1.ListItems.Count + 1, , 2)
                list.SubItems(1) = rs!Bulan
                list.SubItems(2) = rs!Tahun
                list.SubItems(3) = rs!Perusahaan
                list.SubItems(4) = rs!Jabatan
                list.SubItems(5) = Format(Val(rs!GajiYangDibayar), "Rp ###,##,0")
                list.SubItems(6) = Format(Val(rs!PEN_Total_Pendapatan), "Rp ###,##,0")
                list.SubItems(7) = Format(Val(rs!POT_Total_Potongan), "Rp ###,##,0")
                list.SubItems(8) = Format(Val(rs!PEN_Gaji_Pokok_Bulanan), "Rp ###,##,0")
                list.SubItems(9) = Format(Val(rs!PEN_Uang_Makan), "Rp ###,##,0")
                list.SubItems(10) = Format(Val(rs!PEN_Uang_Transportasi), "Rp ###,##,0")
                list.SubItems(11) = Format(Val(rs!PEN_Uang_Lembur), "Rp ###,##,0")
                list.SubItems(12) = Format(Val(rs!PEN_Insentif_Harian), "Rp ###,##,0")
                list.SubItems(13) = Format(Val(rs!PEN_Insentif), "Rp ###,##,0")
                list.SubItems(14) = Format(Val(rs!PEN_JHT), "Rp ###,##,0")
                list.SubItems(15) = Format(Val(rs!PEN_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(16) = Format(Val(rs!PEN_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(17) = Format(Val(rs!PEN_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(18) = Format(Val(rs!PEN_Lain), "Rp ###,##,0")
                list.SubItems(19) = Format(Val(rs!POT_Uang_Makan), "Rp ###,##,0")
                list.SubItems(20) = Format(Val(rs!POT_Absen), "Rp ###,##,0")
                list.SubItems(21) = Format(Val(rs!POT_JHT), "Rp ###,##,0")
                list.SubItems(22) = Format(Val(rs!POT_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(23) = Format(Val(rs!POT_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(24) = Format(Val(rs!POT_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(25) = Format(Val(rs!POT_Lain), "Rp ###,##,0")
                list.SubItems(26) = Format(rs!Tanggal, "YYYY-MM-DD HH:MM")
            rs.MoveNext
        Loop
    
    If FrmGaji.ListView1.ListItems.Count = 0 Then
        MsgBox "Data tidak ada !", vbCritical, "Cari Data"
        FrmGaji.CmdTampilkan.Enabled = True
    End If
End Sub

Sub Cari_Eng()
    FrmGaji.ListView1.ListItems.Clear
    dbConnection
        rs.Open "Select * From vw_Gaji_Juan_Eng" & vbCrLf & _
                "Where " & FrmGaji.CboCari.Text & " Like '%" & FrmGaji.TxtCari.Text & "%'", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            Set list = FrmGaji.ListView1.ListItems.Add(, , FrmGaji.ListView1.ListItems.Count + 1, , 2)
                list.SubItems(1) = rs!Month
                list.SubItems(2) = rs!Year
                list.SubItems(3) = rs!Company
                list.SubItems(4) = rs!Position
                list.SubItems(5) = Format(Val(rs!GajiYangDibayar), "Rp ###,##,0")
                list.SubItems(6) = Format(Val(rs!PEN_Total_Pendapatan), "Rp ###,##,0")
                list.SubItems(7) = Format(Val(rs!POT_Total_Potongan), "Rp ###,##,0")
                list.SubItems(8) = Format(Val(rs!PEN_Gaji_Pokok_Bulanan), "Rp ###,##,0")
                list.SubItems(9) = Format(Val(rs!PEN_Uang_Makan), "Rp ###,##,0")
                list.SubItems(10) = Format(Val(rs!PEN_Uang_Transportasi), "Rp ###,##,0")
                list.SubItems(11) = Format(Val(rs!PEN_Uang_Lembur), "Rp ###,##,0")
                list.SubItems(12) = Format(Val(rs!PEN_Insentif_Harian), "Rp ###,##,0")
                list.SubItems(13) = Format(Val(rs!PEN_Insentif), "Rp ###,##,0")
                list.SubItems(14) = Format(Val(rs!PEN_JHT), "Rp ###,##,0")
                list.SubItems(15) = Format(Val(rs!PEN_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(16) = Format(Val(rs!PEN_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(17) = Format(Val(rs!PEN_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(18) = Format(Val(rs!PEN_Lain), "Rp ###,##,0")
                list.SubItems(19) = Format(Val(rs!POT_Uang_Makan), "Rp ###,##,0")
                list.SubItems(20) = Format(Val(rs!POT_Absen), "Rp ###,##,0")
                list.SubItems(21) = Format(Val(rs!POT_JHT), "Rp ###,##,0")
                list.SubItems(22) = Format(Val(rs!POT_BPJS_Kesehatan), "Rp ###,##,0")
                list.SubItems(23) = Format(Val(rs!POT_Jaminan_Pensiun), "Rp ###,##,0")
                list.SubItems(24) = Format(Val(rs!POT_Pajak_PPh21), "Rp ###,##,0")
                list.SubItems(25) = Format(Val(rs!POT_Lain), "Rp ###,##,0")
                list.SubItems(26) = Format(rs!Tanggal, "YYYY-MM-DD HH:MM")
            rs.MoveNext
        Loop
    
    If FrmGaji.ListView1.ListItems.Count = 0 Then
        MsgBox "Data not found !", vbCritical, "Search Data"
        FrmGaji.CmdTampilkan.Enabled = True
    End If
End Sub

Sub Bersih_Cari()
    FrmGaji.FrmCari.Visible = False
    FrmGaji.TxtCari.Text = ""
    FrmGaji.CmdBatal.Enabled = False
End Sub

Sub Load_Waktu()
    dbConnection
        rs.Open "Select * From APLIKASI.dbo.Detail_Application Where [App Name] = 'List Gaji.exe'", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            If FrmGaji.CboBahasa.Text = "English" Then
                FrmGaji.Caption = "Salary Report During Work  ~  (Last Login : " & Format(rs![Update Time], "DD MMMM YYYY - HH:MM AM/PM") & ")"
            ElseIf FrmGaji.CboBahasa.Text = "Indonesia" Then
                FrmGaji.Caption = "Laporan Gaji Selama Bekerja  ~  (Login Terakhir : " & Format(rs![Update Time], "DD MMMM YYYY - HH:MM") & ")"
            End If
        rs.MoveNext
    Loop
End Sub

Sub Keluar()
    DoEvents
    For I = 1 To FrmGaji.Height
        If FrmGaji.Height < Screen.Width Then
            FrmGaji.Height = Trim(Str(Int(FrmGaji.Height) - 60))
            FrmGaji.Width = Trim(Str(Int(FrmGaji.Width) - 100))
        End If
    Next I

    For I = 1 To FrmGaji.Top
        If FrmGaji.Top < Screen.Width Then
            FrmGaji.Top = Trim(Str(Int(FrmGaji.Top) - 50))
        End If
    Next I
    
    dbConnection
        UpdSQL = "Update APLIKASI.dbo.Detail_Application" & vbCrLf & _
                 "Set [Update Time] = getdate()" & vbCrLf & _
                 "Where [App Name] = 'List Gaji.exe'"
    connect.Execute UpdSQL
    
    End
End Sub

'-- Arsip, udh gak dipake --
Sub ID_Bulan()
    If FrmInput.TxtBulan.Text = "Januari" Then
        KodeBulan = "01"
    ElseIf FrmInput.TxtBulan.Text = "Februari" Then
        KodeBulan = "02"
    ElseIf FrmInput.TxtBulan.Text = "Maret" Then
        KodeBulan = "03"
    ElseIf FrmInput.TxtBulan.Text = "April" Then
        KodeBulan = "04"
    ElseIf FrmInput.TxtBulan.Text = "Mei" Then
        KodeBulan = "05"
    ElseIf FrmInput.TxtBulan.Text = "Juni" Then
        KodeBulan = "06"
    ElseIf FrmInput.TxtBulan.Text = "Juli" Then
        KodeBulan = "07"
    ElseIf FrmInput.TxtBulan.Text = "Agustus" Then
        KodeBulan = "08"
    ElseIf FrmInput.TxtBulan.Text = "September" Then
        KodeBulan = "09"
    ElseIf FrmInput.TxtBulan.Text = "Oktober" Then
        KodeBulan = "10"
    ElseIf FrmInput.TxtBulan.Text = "November" Then
        KodeBulan = "11"
    ElseIf FrmInput.TxtBulan.Text = "Desember" Then
        KodeBulan = "12"
    End If
End Sub
'-- End Arsip --

Sub ComboBox()
    FrmInput.LstBulan.Clear
    dbConnection
    rs.Open "Select * From M_Bulan Order By ID_Bulan Asc", connect, adOpenDynamic, adLockOptimistic
    Do Until rs.EOF
        FrmInput.LstBulan.AddItem rs!Bulan
    rs.MoveNext
    Loop
    
    FrmInput.LstPerusahaan.Clear
    dbConnection
    rs.Open "Select * From M_Pekerjaan Where Status = 1 Order By TanggalMasuk Asc", connect, adOpenDynamic, adLockOptimistic
    Do Until rs.EOF
        FrmInput.LstPerusahaan.AddItem rs!Perusahaan
    rs.MoveNext
    Loop
End Sub

Sub StatusKry()
    dbConnection
    rs.Open "Select * From M_Pekerjaan Where Perusahaan = '" & FrmInput.LstPerusahaan.Text & "'", connect, adOpenDynamic, adLockOptimistic
    Do Until rs.EOF
        FrmInput.LblJabatan.Caption = "[" & rs!Perusahaan & "]" & " - " & "[" & rs!Jabatan & "]"
    rs.MoveNext
    Loop
End Sub

Sub Pekerjaan()
    FrmInput.LvwKry.ListItems.Clear
    dbConnection
        rs.Open "Select NIK, Perusahaan, Bagian, Jabatan, TanggalMasuk, TanggalKeluar," & vbCrLf & _
                "   (Case When Status = 1 Then 'Aktif' Else 'Tidak Aktif' End) As Status," & vbCrLf & _
                "   (Case When JHT = 1 Then 'Ya' Else 'Tidak' End) As JHT," & vbCrLf & _
                "   (Case When JKN = 1 Then 'Ya' Else 'Tidak' End) As JKN," & vbCrLf & _
                "   (Case When NPWP = 1 Then 'Ya' Else 'Tidak' End) As NPWP, Alamat" & vbCrLf & _
                "From M_Pekerjaan Order By TanggalMasuk Asc", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            Set list = FrmInput.LvwKry.ListItems.Add(, , FrmInput.LvwKry.ListItems.Count + 1)
                list.SubItems(1) = rs!NIK
                list.SubItems(2) = rs!Perusahaan
                list.SubItems(3) = rs!Bagian
                list.SubItems(4) = rs!Jabatan
                list.SubItems(5) = Format(rs!TanggalMasuk, "YYYY-MM-DD")
                list.SubItems(6) = Format(rs!TanggalKeluar, "YYYY-MM-DD")
                list.SubItems(7) = rs!Status
                list.SubItems(8) = rs!JHT
                list.SubItems(9) = rs!JKN
                list.SubItems(10) = rs!NPWP
                list.SubItems(11) = rs!Alamat
            rs.MoveNext
        Loop
End Sub

Sub SimpanData()
    dbConnection
        InsSQL = "Insert Into T_Gaji_" & FrmInput.LblTahun.Caption & " "
        InsSQL = InsSQL & "Select ID_Bulan, KodePerusahaan, "
        InsSQL = InsSQL & "'" & FrmInput.TxtGaji.Text & "', getdate() "
        InsSQL = InsSQL & "From M_Pekerjaan "
        InsSQL = InsSQL & "Inner Join M_Bulan"
        InsSQL = InsSQL & "     On Bulan = '" & FrmInput.TxtBulan.Text & "' "
        InsSQL = InsSQL & "Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "'"
    connect.Execute InsSQL
    
    dbConnection
        InsSQL = "Insert Into T_Pendapatan_Gaji_" & FrmInput.LblTahun.Caption & " "
        InsSQL = InsSQL & "Select ID_Bulan, KodePerusahaan, "
        InsSQL = InsSQL & "'" & FrmInput.TxtPenGapok.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenMakan.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenTransport.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenLembur.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenInsHarian.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenIns.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenJHT.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenJKN.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenPensiun.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenPajak.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenLain.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPenTotal.Text & "' "
        InsSQL = InsSQL & "From M_Pekerjaan "
        InsSQL = InsSQL & "Inner Join M_Bulan"
        InsSQL = InsSQL & "   On Bulan = '" & FrmInput.TxtBulan.Text & "' "
        InsSQL = InsSQL & "Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "'"
    connect.Execute InsSQL
    
    dbConnection
        InsSQL = "Insert Into T_Potongan_Gaji_" & FrmInput.LblTahun.Caption & " "
        InsSQL = InsSQL & "Select ID_Bulan, KodePerusahaan, "
        InsSQL = InsSQL & "'" & FrmInput.TxtPotMakan.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotJHT.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotJKN.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotPensiun.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotAbsen.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotPajak.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotLain.Text & "',"
        InsSQL = InsSQL & "'" & FrmInput.TxtPotTotal.Text & "' "
        InsSQL = InsSQL & "From M_Pekerjaan "
        InsSQL = InsSQL & "Inner Join M_Bulan"
        InsSQL = InsSQL & "   On Bulan = '" & FrmInput.TxtBulan.Text & "' "
        InsSQL = InsSQL & "Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "'"
    connect.Execute InsSQL
End Sub

Sub HapusData()
    dbConnection
        DelSQL = "Delete T_Gaji_" & FrmInput.LblTahun.Caption & " "
        DelSQL = DelSQL & "Where Bulan In (Select ID_Bulan From M_Bulan"
        DelSQL = DelSQL & "                 Where Bulan = '" & FrmInput.TxtBulan.Text & "')"
        DelSQL = DelSQL & "And Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        DelSQL = DelSQL & "                     Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "')"
    connect.Execute DelSQL
    
    dbConnection
        DelSQL = "Delete T_Pendapatan_Gaji_" & FrmInput.LblTahun.Caption & " "
        DelSQL = DelSQL & "Where PEN_ID_Bulan In (Select ID_Bulan From M_Bulan"
        DelSQL = DelSQL & "                         Where Bulan = '" & FrmInput.TxtBulan.Text & "')"
        DelSQL = DelSQL & "And PEN_ID_Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        DelSQL = DelSQL & "                             Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "')"
    connect.Execute DelSQL
    
    dbConnection
        DelSQL = "Delete T_Potongan_Gaji_" & FrmInput.LblTahun.Caption & " "
        DelSQL = DelSQL & "Where POT_ID_Bulan In (Select ID_Bulan From M_Bulan"
        DelSQL = DelSQL & "                         Where Bulan = '" & FrmInput.TxtBulan.Text & "')"
        DelSQL = DelSQL & "And POT_ID_Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        DelSQL = DelSQL & "                             Where Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "')"
    connect.Execute DelSQL
End Sub

Sub RubahData()
    dbConnection
        UpdSQL = "Update T_Gaji_" & FrmInput.LblTahun.Caption & " "
        UpdSQL = UpdSQL & "Set Bulan = ID_Bulan, Perusahaan = KodePerusahaan, "
        UpdSQL = UpdSQL & "GajiYangDIbayar = '" & FrmInput.TxtGaji.Text & "', Tanggal = getdate() "
        UpdSQL = UpdSQL & "From T_Gaji_" & FrmInput.LblTahun.Caption & " a "
        UpdSQL = UpdSQL & "Inner Join M_Pekerjaan b"
        UpdSQL = UpdSQL & "    On b.Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "' "
        UpdSQL = UpdSQL & "Inner Join M_Bulan c"
        UpdSQL = UpdSQL & "    On c.Bulan = '" & FrmInput.TxtBulan.Text & "' "
        UpdSQL = UpdSQL & "Where a.Bulan In (Select ID_Bulan From M_Bulan "
        UpdSQL = UpdSQL & "                     Where Bulan = '" & FrmGaji.ListView1.SelectedItem.SubItems(1) & "') "
        UpdSQL = UpdSQL & "And a.Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        UpdSQL = UpdSQL & "                         Where Perusahaan = '" & FrmGaji.ListView1.SelectedItem.SubItems(3) & "')"
    connect.Execute UpdSQL
    
    dbConnection
        UpdSQL = "Update T_Pendapatan_Gaji_" & FrmInput.LblTahun.Caption & " "
        UpdSQL = UpdSQL & "Set PEN_ID_Bulan = ID_Bulan, PEN_ID_Perusahaan = KodePerusahaan, "
        UpdSQL = UpdSQL & "PEN_Gaji_Pokok_Bulanan = '" & FrmInput.TxtPenGapok.Text & "', "
        UpdSQL = UpdSQL & "PEN_Uang_Makan = '" & FrmInput.TxtPenMakan.Text & "', "
        UpdSQL = UpdSQL & "PEN_Uang_Transportasi = '" & FrmInput.TxtPenTransport.Text & "', "
        UpdSQL = UpdSQL & "PEN_Uang_Lembur = '" & FrmInput.TxtPenLembur.Text & "', "
        UpdSQL = UpdSQL & "PEN_Insentif_Harian = '" & FrmInput.TxtPenInsHarian.Text & "', "
        UpdSQL = UpdSQL & "PEN_Insentif = '" & FrmInput.TxtPenIns.Text & "', "
        UpdSQL = UpdSQL & "PEN_JHT = '" & FrmInput.TxtPenJHT.Text & "', "
        UpdSQL = UpdSQL & "PEN_BPJS_Kesehatan = '" & FrmInput.TxtPenJKN.Text & "', "
        UpdSQL = UpdSQL & "PEN_Jaminan_Pensiun = '" & FrmInput.TxtPenPensiun.Text & "', "
        UpdSQL = UpdSQL & "PEN_Pajak_PPh21 = '" & FrmInput.TxtPenPajak.Text & "', "
        UpdSQL = UpdSQL & "PEN_Lain = '" & FrmInput.TxtPenLain.Text & "', "
        UpdSQL = UpdSQL & "PEN_Total_Pendapatan = '" & FrmInput.TxtPenTotal.Text & "' "
        UpdSQL = UpdSQL & "From T_Pendapatan_Gaji_" & FrmInput.LblTahun.Caption & " a "
        UpdSQL = UpdSQL & "Inner Join M_Pekerjaan b"
        UpdSQL = UpdSQL & "    On b.Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "' "
        UpdSQL = UpdSQL & "Inner Join M_Bulan c"
        UpdSQL = UpdSQL & "    On c.Bulan = '" & FrmInput.TxtBulan.Text & "' "
        UpdSQL = UpdSQL & "Where PEN_ID_Bulan In (Select ID_Bulan From M_Bulan "
        UpdSQL = UpdSQL & "                         Where Bulan = '" & FrmGaji.ListView1.SelectedItem.SubItems(1) & "') "
        UpdSQL = UpdSQL & "And PEN_ID_Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        UpdSQL = UpdSQL & "                             Where Perusahaan = '" & FrmGaji.ListView1.SelectedItem.SubItems(3) & "')"
    connect.Execute UpdSQL
    
    dbConnection
        UpdSQL = "Update T_Potongan_Gaji_" & FrmInput.LblTahun.Caption & " "
        UpdSQL = UpdSQL & "Set POT_ID_Bulan = ID_Bulan, POT_ID_Perusahaan = KodePerusahaan, "
        UpdSQL = UpdSQL & "POT_Uang_Makan = '" & FrmInput.TxtPotMakan.Text & "', "
        UpdSQL = UpdSQL & "POT_JHT = '" & FrmInput.TxtPotJHT.Text & "', "
        UpdSQL = UpdSQL & "POT_BPJS_Kesehatan = '" & FrmInput.TxtPotJKN.Text & "', "
        UpdSQL = UpdSQL & "POT_Jaminan_Pensiun = '" & FrmInput.TxtPotPensiun.Text & "', "
        UpdSQL = UpdSQL & "POT_Absen = '" & FrmInput.TxtPotAbsen.Text & "', "
        UpdSQL = UpdSQL & "POT_Pajak_PPh21 = '" & FrmInput.TxtPotPajak.Text & "', "
        UpdSQL = UpdSQL & "POT_Lain = '" & FrmInput.TxtPotLain.Text & "', "
        UpdSQL = UpdSQL & "POT_Total_Potongan = '" & FrmInput.TxtPotTotal.Text & "' "
        UpdSQL = UpdSQL & "From T_Potongan_Gaji_" & FrmInput.LblTahun.Caption & " a "
        UpdSQL = UpdSQL & "Inner Join M_Pekerjaan b"
        UpdSQL = UpdSQL & "    On b.Perusahaan = '" & FrmInput.TxtPerusahaan.Text & "' "
        UpdSQL = UpdSQL & "Inner Join M_Bulan c"
        UpdSQL = UpdSQL & "    On c.Bulan = '" & FrmInput.TxtBulan.Text & "' "
        UpdSQL = UpdSQL & "Where POT_ID_Bulan In (Select ID_Bulan From M_Bulan "
        UpdSQL = UpdSQL & "                         Where Bulan = '" & FrmGaji.ListView1.SelectedItem.SubItems(1) & "') "
        UpdSQL = UpdSQL & "And POT_ID_Perusahaan In (Select KodePerusahaan From M_Pekerjaan"
        UpdSQL = UpdSQL & "                             Where Perusahaan = '" & FrmGaji.ListView1.SelectedItem.SubItems(3) & "')"
    connect.Execute UpdSQL
End Sub

Sub TampilRP()
    dbConnection
        rs.Open "Select * From vw_Gaji_Juan" & vbCrLf & _
                "Where Bulan = '" & FrmGaji.ListView1.SelectedItem.SubItems(1) & "'" & vbCrLf & _
                "And Tahun = '" & FrmGaji.ListView1.SelectedItem.SubItems(2) & "'" & vbCrLf & _
                "And Perusahaan = '" & FrmGaji.ListView1.SelectedItem.SubItems(3) & "'", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            FrmInput.TxtBulan.Text = rs!Bulan
            FrmInput.LblTahun.Caption = rs!Tahun
            FrmInput.TxtPerusahaan.Text = rs!Perusahaan
            FrmInput.LblJabatan.Caption = "[" & rs!Perusahaan & "] - [" & rs!Jabatan & "]"
            FrmInput.TxtGaji.Text = Format(Val(rs!GajiYangDibayar), "Rp ###,##,0")
            FrmInput.TxtPenTotal.Text = Format(Val(rs!PEN_Total_Pendapatan), "Rp ###,##,0")
            FrmInput.TxtPotTotal.Text = Format(Val(rs!POT_Total_Potongan), "Rp ###,##,0")
            FrmInput.TxtPenGapok.Text = Format(Val(rs!PEN_Gaji_Pokok_Bulanan), "Rp ###,##,0")
            FrmInput.TxtPenMakan.Text = Format(Val(rs!PEN_Uang_Makan), "Rp ###,##,0")
            FrmInput.TxtPenTransport.Text = Format(Val(rs!PEN_Uang_Transportasi), "Rp ###,##,0")
            FrmInput.TxtPenLembur.Text = Format(Val(rs!PEN_Uang_Lembur), "Rp ###,##,0")
            FrmInput.TxtPenInsHarian.Text = Format(Val(rs!PEN_Insentif_Harian), "Rp ###,##,0")
            FrmInput.TxtPenIns.Text = Format(Val(rs!PEN_Insentif), "Rp ###,##,0")
            FrmInput.TxtPenJHT.Text = Format(Val(rs!PEN_JHT), "Rp ###,##,0")
            FrmInput.TxtPenJKN.Text = Format(Val(rs!PEN_BPJS_Kesehatan), "Rp ###,##,0")
            FrmInput.TxtPenPensiun.Text = Format(Val(rs!PEN_Jaminan_Pensiun), "Rp ###,##,0")
            FrmInput.TxtPenPajak.Text = Format(Val(rs!PEN_Pajak_PPh21), "Rp ###,##,0")
            FrmInput.TxtPenLain.Text = Format(Val(rs!PEN_Lain), "Rp ###,##,0")
            FrmInput.TxtPotMakan.Text = Format(Val(rs!POT_Uang_Makan), "Rp ###,##,0")
            FrmInput.TxtPotAbsen.Text = Format(Val(rs!POT_Absen), "Rp ###,##,0")
            FrmInput.TxtPotJHT.Text = Format(Val(rs!POT_JHT), "Rp ###,##,0")
            FrmInput.TxtPotJKN.Text = Format(Val(rs!POT_BPJS_Kesehatan), "Rp ###,##,0")
            FrmInput.TxtPotPensiun.Text = Format(Val(rs!POT_Jaminan_Pensiun), "Rp ###,##,0")
            FrmInput.TxtPotPajak.Text = Format(Val(rs!POT_Pajak_PPh21), "Rp ###,##,0")
            FrmInput.TxtPotLain.Text = Format(Val(rs!POT_Lain), "Rp ###,##,0")
        rs.MoveNext
    Loop
End Sub

Sub TampilNotRP()
    dbConnection
        rs.Open "Select * From vw_Gaji_Juan" & vbCrLf & _
                "Where Bulan = '" & FrmGaji.ListView1.SelectedItem.SubItems(1) & "'" & vbCrLf & _
                "And Tahun = '" & FrmGaji.ListView1.SelectedItem.SubItems(2) & "'" & vbCrLf & _
                "And Perusahaan = '" & FrmGaji.ListView1.SelectedItem.SubItems(3) & "'", connect, adOpenDynamic, adLockOptimistic
        Do Until rs.EOF
            FrmInput.TxtBulan.Text = rs!Bulan
            FrmInput.LblTahun.Caption = rs!Tahun
            FrmInput.TxtPerusahaan.Text = rs!Perusahaan
            FrmInput.LblJabatan.Caption = "[" & rs!Perusahaan & "] - [" & rs!Jabatan & "]"
            FrmInput.TxtGaji.Text = rs!GajiYangDibayar
            FrmInput.TxtPenTotal.Text = rs!PEN_Total_Pendapatan
            FrmInput.TxtPotTotal.Text = rs!POT_Total_Potongan
            FrmInput.TxtPenGapok.Text = rs!PEN_Gaji_Pokok_Bulanan
            FrmInput.TxtPenMakan.Text = rs!PEN_Uang_Makan
            FrmInput.TxtPenTransport.Text = rs!PEN_Uang_Transportasi
            FrmInput.TxtPenLembur.Text = rs!PEN_Uang_Lembur
            FrmInput.TxtPenInsHarian.Text = rs!PEN_Insentif_Harian
            FrmInput.TxtPenIns.Text = rs!PEN_Insentif
            FrmInput.TxtPenJHT.Text = rs!PEN_JHT
            FrmInput.TxtPenJKN.Text = rs!PEN_BPJS_Kesehatan
            FrmInput.TxtPenPensiun.Text = rs!PEN_Jaminan_Pensiun
            FrmInput.TxtPenPajak.Text = rs!PEN_Pajak_PPh21
            FrmInput.TxtPenLain.Text = rs!PEN_Lain
            FrmInput.TxtPotMakan.Text = rs!POT_Uang_Makan
            FrmInput.TxtPotAbsen.Text = rs!POT_Absen
            FrmInput.TxtPotJHT.Text = rs!POT_JHT
            FrmInput.TxtPotJKN.Text = rs!POT_BPJS_Kesehatan
            FrmInput.TxtPotPensiun.Text = rs!POT_Jaminan_Pensiun
            FrmInput.TxtPotPajak.Text = rs!POT_Pajak_PPh21
            FrmInput.TxtPotLain.Text = rs!POT_Lain
        rs.MoveNext
    Loop
End Sub

Sub M_Pekerjaan()
    dbConnection
    rs.Open "Select * From M_Pekerjaan Where Perusahaan = '" & FrmInput.LvwKry.SelectedItem.SubItems(2) & "'", connect, adOpenDynamic, adLockOptimistic
    Do Until rs.EOF
    
    Dim varChcStatus, varChcJHT, varChcJKN, varChcNPWP As Integer
        varChcStatus = rs!Status
        varChcJHT = rs!JHT
        varChcJKN = rs!JKN
        varChcNPWP = rs!NPWP
        
        If varChcStatus = True Then
            varChcStatus = 1
        ElseIf varChcStatus = False Then
            varChcStatus = 0
        End If
            
        If varChcJHT = True Then
            varChcJHT = 1
        ElseIf varChcJHT = False Then
            varChcJHT = 0
        End If
        
        If varChcJKN = True Then
            varChcJKN = 1
        ElseIf varChcJKN = False Then
            varChcJKN = 0
        End If
        
        If varChcNPWP = True Then
            varChcNPWP = 1
        ElseIf varChcNPWP = False Then
            varChcNPWP = 0
        End If

        FrmPekerjaan.TxtNo.Text = rs!No
        FrmPekerjaan.TxtNIK.Text = rs!NIK
        FrmPekerjaan.TxtKode.Text = rs!KodePerusahaan
        FrmPekerjaan.TxtPerusahaan.Text = rs!Perusahaan
        FrmPekerjaan.TxtBagian.Text = rs!Bagian
        FrmPekerjaan.TxtJabatan.Text = rs!Jabatan
        FrmPekerjaan.DTMasuk.Value = rs!TanggalMasuk
        FrmPekerjaan.ChcStatus.Value = varChcStatus
        FrmPekerjaan.DTKeluar.Value = rs!TanggalKeluar
        FrmPekerjaan.ChcJHT.Value = varChcJHT
        FrmPekerjaan.ChcJKN.Value = varChcJKN
        FrmPekerjaan.ChcNPWP.Value = varChcNPWP
        FrmPekerjaan.TxtAlamat.Text = rs!Alamat
        FrmPekerjaan.LblTanggal.Caption = "Waktu Update : " & Format(rs!Tanggal, "YYYY-MM-DD HH:MM")
    rs.MoveNext
    Loop
End Sub

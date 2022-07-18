﻿Imports System.Data.OleDb
Public Class frmJurnalUmum

    'Menciptakan objek bernama objJurnal
    Dim objJurnal As New clsJurnalUmum

    Public mPosted As String
    Public mDK As String
    Dim mNoTransaksi As String
    Dim mJumlah As Long

    Sub KondisiAwal()
        Tutup()

        txtKeterangan.Text = ""
        txtTgl.Text = Now
        txtDebet.Text = "0"
        txtKredit.Text = "0"
        lblDebet.Text = "0"
        lblKredit.Text = "0"
        ListView.Clear()

        CariPeriode()

        PosisiListGrid()
        IsiListGridDJurnal()
        NoTransaksi()

        cmdTambah.Text = "&New Item"
        cmdKeluar.Text = "&Keluar"
        cmdTambah.Enabled = True
        cmdSimpan.Enabled = False
        cmdEdit.Enabled = False
        cmdHapus.Enabled = False
    End Sub


    Public Sub PosisiListGrid()
        With ListView.Columns
            .Add("NO.Transaksi", 0)
            .Add("No.Rek", 68)
            .Add("Nama Perkiraan", 360)
            .Add("DK", 0)
            .Add("Debet", 125, HorizontalAlignment.Right)
            .Add("Kredit", 125, HorizontalAlignment.Right)
        End With
    End Sub

    Public Sub IsiListGridDJurnal()
        Try
            Query = "SELECT dJurnal.NoTransaksi, dJurnal.NoPerkiraan, tblMasterPerkiraan.NamaPerkiraan, dJurnal.DK, dJurnal.Debet, dJurnal.Kredit FROM (dJurnal LEFT JOIN hJurnal ON dJurnal.NoTransaksi = hJurnal.NoTransaksi) LEFT JOIN tblMasterPerkiraan ON dJurnal.NoPerkiraan = tblMasterPerkiraan.NoPerkiraan WHERE(((dJurnal.NoTransaksi) = '" & txtNoTransaksi.Text & "'))"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            ListView.Items.Clear()
            For a = 0 To dsData.Tables(0).Rows.Count - 1
                With ListView
                    .Items.Add(dsData.Tables(0).Rows(a).Item(0))
                    .Items(a).SubItems.Add(dsData.Tables(0).Rows(a).Item(1))
                    .Items(a).SubItems.Add(dsData.Tables(0).Rows(a).Item(2))
                    .Items(a).SubItems.Add(dsData.Tables(0).Rows(a).Item(3))
                    .Items(a).SubItems.Add(Format(dsData.Tables(0).Rows(a).Item(4), "###,###"))
                    .Items(a).SubItems.Add(Format(dsData.Tables(0).Rows(a).Item(5), "###,###"))
                    If (a Mod 2 = 0) Then
                        .Items(a).BackColor = Color.LightSteelBlue
                    Else
                        .Items(a).BackColor = Color.LightBlue
                    End If
                End With
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NoTransaksi()
        Try
            Dim mMemoriNo As String
            Dim mTempNoTransaksi As String

            mMemoriNo = Format(Microsoft.VisualBasic.Right(Now.Year, 2)) & Format(Now.Month, "00") & Format(Now.Day, "00")
            txtNoTransaksi.Text = "JU-" & mMemoriNo & "000"
            Query = "SELECT Count(hJurnal.NoTransaksi) AS CountOfNoTransaksi FROM(hJurnal) HAVING ((Mid([hJurnal].[NoTransaksi],4,6)='" & mMemoriNo & "'))"

            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count = 0 Then
                txtNoTransaksi.Text = "JU-" & mMemoriNo + "000" + dsData.Tables(0).Rows(0).Item(0) + 1
                mTempNoTransaksi = "JU-" & mMemoriNo + "000" + dsData.Tables(0).Rows(0).Item(0) + 1
            Else
                txtNoTransaksi.Text = "JU-" & mMemoriNo + "000" + dsData.Tables(0).Rows(0).Item(0) + 1
                mTempNoTransaksi = "JU-" & mMemoriNo + "000" + dsData.Tables(0).Rows(0).Item(0) + 1
            End If

            Query = "SELECT hJurnal.NoTransaksi, Count(hJurnal.NoTransaksi) AS Jumlah FROM hJurnal GROUP BY hJurnal.NoTransaksi HAVING (((hJurnal.NoTransaksi)='" & txtNoTransaksi.Text & "'))"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count = 0 Then
                txtNoTransaksi.Text = mTempNoTransaksi
            Else
                txtNoTransaksi.Text = Microsoft.VisualBasic.Mid(txtNoTransaksi.Text, 1, 3) & Val(Microsoft.VisualBasic.Mid(txtNoTransaksi.Text, 4, 9)) + 1
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub TotalDebetKredit()
        Try
            QueryDebetKredit = "SELECT dJurnal.NoTransaksi, Sum(dJurnal.Debet) AS TotalDebet, Sum(dJurnal.Kredit) AS TotalKredit FROM(dJurnal) GROUP BY dJurnal.NoTransaksi HAVING (((dJurnal.NoTransaksi)= '" & txtNoTransaksi.Text & "'))"
            daDataDebetKredit = New OleDbDataAdapter(QueryDebetKredit, conn)
            dsDataDebetKredit = New DataSet
            daDataDebetKredit.Fill(dsDataDebetKredit)

            With dsDataDebetKredit.Tables(0).Rows(0)
                lblDebet.Text = Format(.Item(1), "#,#")
                lblKredit.Text = Format(.Item(2), "#,#")
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub AmbilData()
        Try
            With ListView.SelectedItems
                mNoTransaksi = .Item(0).SubItems(0).Text
                txtNoPerkiraan.Text = .Item(0).SubItems(1).Text
                lblNamaPerkiraan.Text = .Item(0).SubItems(2).Text
                mDK = .Item(0).SubItems(3).Text
                If .Item(0).SubItems(4).Text = "" Then
                    txtDebet.Text = 0
                Else
                    txtDebet.Text = .Item(0).SubItems(4).Text
                End If

                If .Item(0).SubItems(5).Text = "" Then
                    txtKredit.Text = 0
                Else
                    txtKredit.Text = .Item(0).SubItems(5).Text
                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub HapusIsiGrid()
        Try
            'Menghapus isi data grid
            Query = "DELETE FROM dJurnal WHERE NoTransaksi = '" & txtNoTransaksi.Text & "' AND NoPerkiraan = '" & txtNoPerkiraan.Text & "' AND DK = '" & mDK & "'"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CariDataNoTransaksi()
        Try
            Query = "SELECT hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status FROM(hJurnal) GROUP BY hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status HAVING(((hJurnal.NoTransaksi) = '" & txtNoTransaksi.Text & "')) ORDER BY hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count - 1 Then
                BersihkanIsian()
                txtTgl.Focus()
                ListView.Clear()
                PosisiListGrid()
                IsiListGridDJurnal()
            Else
                lblPeriode.Text = dsData.Tables(0).Rows(0).Item(0)
                txtTgl.Text = dsData.Tables(0).Rows(0).Item(1)
                txtNoTransaksi.Text = dsData.Tables(0).Rows(0).Item(2)
                txtKeterangan.Text = dsData.Tables(0).Rows(0).Item(3)
                mPosted = dsData.Tables(0).Rows(0).Item(4)

                IsiListGridDJurnal()
                TotalDebetKredit()
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Sub rutin untuk menyimpan data ke hJurnal
    Private Sub PeriksaDataNoTransaksi()
        With objJurnal
            Try
                Query = "SELECT hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status FROM(hJurnal) GROUP BY hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status HAVING(((hJurnal.NoTransaksi) = '" & txtNoTransaksi.Text & "')) ORDER BY hJurnal.Periode, hJurnal.TglTransaksi, hJurnal.NoTransaksi, hJurnal.Keterangan, hJurnal.Status"
                daData = New OleDbDataAdapter(Query, conn)
                dsData = New DataSet
                daData.Fill(dsData)

                If dsData.Tables(0).Rows.Count - 1 Then
                    'Menyimpan data ke hJurnal
                    .SimpanDataHJurnal()
                End If
            Catch ex As Exception
            End Try
        End With
    End Sub

    Private Sub AdaNoTransaksi()
        Try
            Query = "SELECT noTransaksi FROM hJurnal WHERE noTransaksi = '" & txtNoTransaksi.Text & "'"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count - 1 Then
                MsgBox("Belum ada transksi jurnal....", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Pesan simpan data")
                txtNoPerkiraan.Focus()
            Else
                PeriksaDataNoTransaksi()
                BersihkanIsian()
                BersihkanIsianGrid()
                NoTransaksi()
                ListView.Clear()
                PosisiListGrid()
                IsiListGridDJurnal()
                cmdEdit.Text = "&Edit"
                cmdTambah.Text = "&New Item"
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CariNoPerkiraan()
        Try
            Query = "SELECT tblMasterPerkiraan.NoPerkiraan, tblMasterPerkiraan.NamaPerkiraan FROM(tblMasterPerkiraan) WHERE (((tblMasterPerkiraan.NoPerkiraan)='" & txtNoPerkiraan.Text & "'))"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count - 1 Then
                MsgBox("No.Perkiraan ini tidak ada", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Cari no perkiraan")
                BersihkanIsianGrid()
                txtNoPerkiraan.Focus()
            Else
                txtNoPerkiraan.Text = dsData.Tables(0).Rows(0).Item(0)
                lblNamaPerkiraan.Text = dsData.Tables(0).Rows(0).Item(1)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EditJurnalGrid()
        With objJurnal
            Try
                .EditDataHJurnal() 'edit hJurnal

                If txtDebet.Text > 0 Then
                    mDK = "D"
                Else
                    mDK = "K"
                End If
                .EditData() 'edit dJurnal
                IsiListGridDJurnal()
                BersihkanIsianGrid()
                txtNoPerkiraan.Enabled = True
                txtNoTransaksi.Enabled = True
                TotalDebetKredit()
            Catch ex As Exception
            End Try
        End With
    End Sub

    Private Sub CariPeriode()
        Try
            Query = "SELECT tblMasterPeriode.Periode FROM(tblMasterPeriode)WHERE(((tblMasterPeriode.Keterangan) = '" & "UnPosted" & "')) ORDER BY tblMasterPeriode.Periode"
            daData = New OleDbDataAdapter(Query, conn)
            dsData = New DataSet
            daData.Fill(dsData)

            If dsData.Tables(0).Rows.Count - 1 Then
                MsgBox("Periode belum ada, silahkan buat periode di master periode...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Pesan")
                Dispose()
            Else
                lblPeriode.Text = dsData.Tables(0).Rows(0).Item(0)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub BersihkanIsian()
        txtTgl.Text = txtTgl.Value
        txtKeterangan.Clear()
        txtNoPerkiraan.Clear()
        lblNamaPerkiraan.Text = ""
        txtDebet.Text = "0"
        txtKredit.Text = "0"
        lblDebet.Text = "0"
        lblKredit.Text = "0"
        txtTgl.Focus()
    End Sub

    Public Sub BersihkanIsianGrid()
        txtNoPerkiraan.Clear()
        lblNamaPerkiraan.Text = ""
        txtDebet.Text = "0"
        txtKredit.Text = "0"
        txtNoPerkiraan.Focus()
    End Sub

    Sub Buka()
        txtNoPerkiraan.Enabled = True
        txtTgl.Enabled = True
        txtKeterangan.Enabled = True
        txtNoPerkiraan.Enabled = True
        txtDebet.Enabled = True
        txtKredit.Enabled = True
    End Sub

    Sub Tutup()
        txtNoPerkiraan.Enabled = False
        txtTgl.Enabled = False
        txtKeterangan.Enabled = False
        txtNoPerkiraan.Enabled = False
        txtDebet.Enabled = False
        txtKredit.Enabled = False
    End Sub

    Private Sub frmJurnalUmum_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            KoneksiKeAccess()

            KondisiAwal()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdKeluar_Click(sender As Object, e As EventArgs) Handles cmdKeluar.Click
        If cmdKeluar.Text = "&Batal" Then
            KondisiAwal()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub cmdSimpan_Click(sender As Object, e As EventArgs) Handles cmdSimpan.Click
        Try
            If txtKeterangan.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "")
                txtKeterangan.Focus()
            Else
                If lblDebet.Text <> lblKredit.Text Then
                    MsgBox("Jumlah debet dan kredit tidak seimbang, silahkan periksa", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
                Else
                    AdaNoTransaksi()
                    KondisiAwal()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdTambah_Click(sender As Object, e As EventArgs) Handles cmdTambah.Click
        Buka()
        Try
            dsData.Reset()
            IsiListGridDJurnal()
            BersihkanIsian()
            txtKeterangan.Focus()
            txtTgl.Text = Now
            NoTransaksi()
            ListView.Clear()
            PosisiListGrid()
            txtNoPerkiraan.Enabled = True
            txtNoTransaksi.Enabled = True
            cmdEdit.Text = "&Edit"
            cmdTambah.Text = "&New Item"
            cmdKeluar.Text = "&Batal"
            cmdSimpan.Enabled = True
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtKredit_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtKredit.KeyPress
        With objJurnal
            If e.KeyChar = Chr(13) Then
                If cmdTambah.Text = "&New Item" Then
                    Try
                        If txtNoPerkiraan.Text = "" Then
                            MessageBox.Show("No.Perkiraan masih kosong", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            txtNoPerkiraan.Focus()
                        Else
                            If cmdEdit.Text = "&Edit" Then
                                If txtDebet.Text > 0 Then
                                    mDK = "D"
                                Else
                                    mDK = "K"
                                End If

                                If txtDebet.Text = "" Then
                                    txtDebet.Text = 0
                                Else
                                    txtDebet.Text = txtDebet.Text
                                End If

                                If txtKredit.Text = "" Then
                                    txtKredit.Text = 0
                                Else
                                    txtKredit.Text = txtKredit.Text
                                End If

                                PeriksaDataNoTransaksi()
                                .SimpanData() 'dJurnal
                                mPosted = "UnPosted"
                                TotalDebetKredit()
                                IsiListGridDJurnal()
                                BersihkanIsianGrid()
                                txtNoPerkiraan.Focus()
                            Else
                                cmdEdit.Text = "&Edit"
                                EditJurnalGrid()
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    If mPosted = "UnPosted" Then
                        Try
                            If txtNoPerkiraan.Text = "" Then
                                MessageBox.Show("No.Perkiraan masih kosong", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                txtNoPerkiraan.Focus()
                            Else
                                If cmdEdit.Text = "&Edit" Then
                                    If txtDebet.Text > 0 Then
                                        mDK = "D"
                                    Else
                                        mDK = "K"
                                    End If

                                    If txtDebet.Text = "" Then
                                        txtDebet.Text = 0
                                    Else
                                        txtDebet.Text = txtDebet.Text
                                    End If

                                    If txtKredit.Text = "" Then
                                        txtKredit.Text = 0
                                    Else
                                        txtKredit.Text = txtKredit.Text
                                    End If

                                    .SimpanData() 'dJurnal
                                    mPosted = "UnPosted"
                                    TotalDebetKredit()

                                    IsiListGridDJurnal()
                                    BersihkanIsianGrid()
                                    txtNoPerkiraan.Focus()
                                Else
                                    cmdEdit.Text = "&Edit"
                                    EditJurnalGrid()
                                End If
                            End If
                        Catch ex As Exception
                            MsgBox("Silahkan tekan tombol tambah data", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Pesan kesalahan")
                            BersihkanIsianGrid()
                        End Try
                    Else
                        MsgBox("Data ini sudah diposting, tidak bisa edit / tambah data", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan Edit Data")
                        txtNoPerkiraan.Focus()
                        BersihkanIsianGrid()
                    End If
                End If
            End If

        End With

        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack Or e.KeyChar = ".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNoPerkiraan_DoubleClick(sender As Object, e As EventArgs) Handles txtNoPerkiraan.DoubleClick
        If txtKeterangan.Text = "" Then
            MsgBox("Keterangan masih kosong, silahkan isi", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
            txtKeterangan.Focus()
        Else
            frmSubNoPerkiraanJurnalUmum.ShowDialog()
            If Len(txtNoPerkiraan.Text) <> 0 Then
                txtDebet.Focus()
            Else
                txtNoPerkiraan.Focus()
            End If
        End If
    End Sub

    Private Sub txtNoPerkiraan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNoPerkiraan.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtKeterangan.Text = "" Then
                MsgBox("Keterangan tidak boleh kosong, silahkan isi", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Isi keterangan")
                txtKeterangan.Focus()
            Else
                If txtNoPerkiraan.Text = "" Then
                    cmdSimpan.Focus()
                Else
                    CariNoPerkiraan()
                    txtDebet.Focus()
                End If
            End If
        End If

        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDebet_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDebet.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtDebet.Text = "" Then
                txtDebet.Text = 0
                txtKredit.Focus()
            Else
                txtKredit.Focus()
            End If
        End If

        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack Or e.KeyChar = ".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNoTransaksi_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNoTransaksi.KeyPress
        If e.KeyChar = Chr(13) Then
            CariDataNoTransaksi()
        End If
    End Sub

    Private Sub txtKeterangan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtKeterangan.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtKeterangan.Text = "" Then
                MsgBox("Keterangan masih kosong, silahkan isi", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
                txtKeterangan.Focus()
            Else
                txtNoPerkiraan.Focus()
            End If
        End If
    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Dim A As String

        With objJurnal
            If mPosted = "UnPosted" Then
                A = MsgBox("Benar akan di-Edit", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "Informasi Edit")
                Select Case A
                    Case vbCancel
                        txtNoTransaksi.Focus()
                        cmdEdit.Text = "&Edit"
                        cmdSimpan.Enabled = True
                        BersihkanIsianGrid()
                        Exit Sub
                    Case vbOK
                        Try
                            If txtNoPerkiraan.Text = "" Then
                                .EditDataHJurnal() 'EditHJurnal
                                TotalDebetKredit()
                                cmdEdit.Text = "&Edit"
                                KondisiAwal()
                            Else
                                EditJurnalGrid()
                                cmdEdit.Text = "&Edit"
                                KondisiAwal()
                            End If
                        Catch ex As Exception
                            MsgBox("Terjadi kesalahan")
                        End Try
                End Select
            Else
                MsgBox("Data ini sudah diposting, tidak bisa di Edit", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Pesan edit")
                cmdEdit.Text = "&Edit"
                KondisiAwal()
                BersihkanIsianGrid()
            End If
        End With
    End Sub

    Private Sub cmdHapus_Click(sender As Object, e As EventArgs) Handles cmdHapus.Click
        Dim A As String

        With objJurnal
            If mPosted = "UnPosted" Then
                Try
                    If Len(txtNoTransaksi.Text) = 0 Then
                        MsgBox("Pilih data yang dihapus", MsgBoxStyle.Information, "Informasi hapus")
                        txtNoPerkiraan.Enabled = True
                        txtNoPerkiraan.Focus()
                        cmdSimpan.Enabled = True
                        Exit Sub
                    Else
                        If txtNoPerkiraan.Text = "" Then
                            'hapus semua transaksi dengan no transksi yang sesuai dengan no.transaksi
                            A = MsgBox("Benar akan dihapus...", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "Informasi")
                            Select Case A
                                Case vbCancel
                                    txtNoPerkiraan.Focus()
                                    cmdEdit.Text = "&Edit"
                                    cmdTambah.Text = "&New Item"
                                    txtNoPerkiraan.Enabled = True
                                    KondisiAwal()
                                    Exit Sub
                                Case vbOK
                                    'Hapus hJurnal
                                    .HapusDataHJurnal()
                                    'Hapus dJurnal
                                    .HapusData()
                                    BersihkanIsian()
                                    BersihkanIsianGrid()
                                    ListView.Clear()
                                    PosisiListGrid()
                                    IsiListGridDJurnal()
                                    txtKeterangan.Focus()
                                    NoTransaksi()
                                    TotalDebetKredit()
                                    cmdEdit.Text = "&Edit"
                                    cmdTambah.Text = "&New Item"
                                    txtNoPerkiraan.Enabled = True
                                    KondisiAwal()
                            End Select
                        Else
                            'untuk menghapus record jurnal
                            A = MsgBox("Benar akan dihapus...", MsgBoxStyle.OkCancel, "Informasi")
                            Select Case A
                                Case vbCancel
                                    txtNoPerkiraan.Focus()
                                    cmdEdit.Text = "&Edit"
                                    cmdTambah.Text = "&New Item"
                                    txtNoPerkiraan.Enabled = True
                                    cmdSimpan.Enabled = True
                                    Exit Sub
                                Case vbOK
                                    HapusIsiGrid()
                                    IsiListGridDJurnal()
                                    BersihkanIsianGrid()
                                    txtNoPerkiraan.Focus()
                                    TotalDebetKredit()
                                    cmdEdit.Text = "&Edit"
                                    cmdTambah.Text = "&New Item"
                                    txtNoPerkiraan.Enabled = True
                                    cmdSimpan.Enabled = True
                            End Select
                        End If
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, "Perhatian")
                End Try
            Else
                MsgBox("Data ini sudah diposting, tidak bisa dihapus...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Hapus data")
                cmdSimpan.Enabled = True
            End If
        End With
    End Sub

    Private Sub ListView_DoubleClick(sender As Object, e As EventArgs) Handles ListView.DoubleClick
        Try
            AmbilData()
            txtNoTransaksi.Enabled = False
            txtNoPerkiraan.Enabled = False
            cmdEdit.Text = "&Update"
            cmdSimpan.Enabled = False
            cmdEdit.Enabled = True
            cmdHapus.Enabled = False
            cmdKeluar.Text = "&Batal"
            Buka()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ListView_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView.KeyPress
        If e.KeyChar = Chr(13) Then
            Try
                AmbilData()
                txtNoTransaksi.Enabled = False
                txtNoPerkiraan.Enabled = False
                cmdEdit.Text = "&Update"
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub ListView_KeyUp(sender As Object, e As KeyEventArgs) Handles ListView.KeyUp
        Dim A As String

        If e.KeyCode = Keys.Escape Then
            BersihkanIsianGrid()
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
            cmdEdit.Text = "&Edit"
            cmdSimpan.Enabled = True
        End If

        If e.KeyCode = Keys.Delete Then
            AmbilData()
            A = MsgBox("Benar akan dihapus...", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "Informasi")
            Select Case A
                Case vbCancel
                    txtNoPerkiraan.Enabled = True
                    cmdSimpan.Enabled = True
                    txtNoPerkiraan.Focus()
                    BersihkanIsianGrid()
                    Exit Sub
                Case vbOK
                    If mPosted = "UnPosted" Then
                        HapusIsiGrid()
                        IsiListGridDJurnal()
                        BersihkanIsianGrid()
                        txtNoPerkiraan.Enabled = True
                        cmdSimpan.Enabled = True
                        txtNoPerkiraan.Focus()
                        TotalDebetKredit()
                    Else
                        MsgBox("Data ini sudah diposting, tidak bisa dihapus", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
                        BersihkanIsianGrid()
                        cmdSimpan.Enabled = True
                    End If
            End Select
        End If
    End Sub

    Private Sub cmdTransaksi_Click(sender As Object, e As EventArgs) Handles cmdTransaksi.Click
        'If lblDebet.Text <> lblKredit.Text Then
        '    MsgBox("Jumlah debet dan kredit tidak seimbang, silahkan periksa", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
        '    txtNoPerkiraan.Enabled = True
        '    txtNoPerkiraan.Focus()
        'Else
        frmSubJurnalUmum.ShowDialog()
        cmdEdit.Text = "&Edit"
        txtTgl.Focus()
        BersihkanIsianGrid()
        cmdSimpan.Enabled = False
        'End If
    End Sub
End Class
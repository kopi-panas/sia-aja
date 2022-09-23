Imports System.Data.OleDb
Public Class frmPerkiraan

    Sub KondisiAwal()
        cmdTambah.Enabled = True
        cmdEdit.Enabled = False
        cmdSimpan.Enabled = False
        cmdHapus.Enabled = False
        cmdKeluar.Enabled = True
        cboGroup.Focus()
        'txtNoPerkiraan.Focus()
        cmdKeluar.Text = "Keluar"
        cmdEdit.Text = "Edit"
        BersihkanIsian()
        Tutup()
    End Sub

    Sub Buka()
        txtNoPerkiraan.Enabled = False
        txtNamaPerkiraan.Enabled = True
        cboKeterangan.Enabled = True
        cboGroup.Enabled = True
    End Sub

    Sub Tutup()
        txtNoPerkiraan.Enabled = False
        txtNamaPerkiraan.Enabled = False
        cboKeterangan.Enabled = False
        cboGroup.Enabled = False
    End Sub

    Public Sub PosisiList()
        With ListView.Columns
            .Add("Kode", 60)
            .Add("Nama Perkiraan", 310)
            .Add("Group", 120)
            .Add("Ket", 50)
        End With
    End Sub

    Public Sub IsiList()
        Try
            Query = "SELECT * FROM tblMasterPerkiraan ORDER BY NoPerkiraan"
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
                    If (a Mod 2 = 0) Then
                        .Items(a).BackColor = Color.Transparent
                    Else
                        .Items(a).BackColor = Color.White
                    End If
                End With
            Next
        Catch ex As Exception
        End Try
    End Sub

    Public Sub AmbilDataDariList()
        Try
            With ListView.SelectedItems
                txtNoPerkiraan.Text = .Item(0).SubItems(0).Text
                txtNamaPerkiraan.Text = .Item(0).SubItems(1).Text
                cboGroup.Text = .Item(0).SubItems(2).Text
                cboKeterangan.Text = .Item(0).SubItems(3).Text
            End With
        Catch ex As Exception
        End Try
    End Sub

    Public Sub BersihkanIsian()
        Try
            txtNoPerkiraan.Clear()
            txtNamaPerkiraan.Clear()
            'txtNoPerkiraan.Enabled = True
            'txtNoPerkiraan.Focus()
            cboGroup.Text = ""
            cboKeterangan.Text = "-"
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtNamaPerkiraan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNamaPerkiraan.KeyPress
        If (e.KeyChar = Chr(13)) Then
            If (txtNamaPerkiraan.Text = "") Then
                MsgBox("Perkiraan tidak boleh kosong", MsgBoxStyle.Critical, "Error")
            Else
                cboGroup.Focus()
            End If
        End If
    End Sub

    Private Sub txtNoPerkiraan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNoPerkiraan.KeyPress
        If e.KeyChar = Chr(13) Then
            If txtNoPerkiraan.Text = "" Then
                MsgBox("NoPerkiraan tidak boleh kosong", MsgBoxStyle.Critical, "Error")
            Else
                If Len(txtNoPerkiraan.Text) < 5 Then
                    MsgBox("Panjang digit kode harus 5 digit", MsgBoxStyle.Exclamation, "Error")
                    txtNoPerkiraan.Focus()
                Else
                    txtNamaPerkiraan.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub txtNoPerkiraan_KeyUp(sender As Object, e As KeyEventArgs) Handles txtNoPerkiraan.KeyUp
        If e.KeyCode = Keys.Escape Then
            BersihkanIsian()
        End If
    End Sub

    Private Sub frmPerkiraan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            KondisiAwal()
            KoneksiKeAccess()
            PosisiList()
            IsiList()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdKeluar_Click(sender As Object, e As EventArgs) Handles cmdKeluar.Click
        If cmdKeluar.Text = "Batal" Then
            KondisiAwal()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub cmdSimpan_Click(sender As Object, e As EventArgs) Handles cmdSimpan.Click
        Try
            If Len(txtNoPerkiraan.Text) < 5 Or txtNoPerkiraan.Text = "" Or txtNamaPerkiraan.Text = "" Or cboGroup.Text = "" Or cboKeterangan.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap atau digit kode akun kurang dari 5!", MsgBoxStyle.Information, "")
            Else
                KoneksiKeAccess()
                Dim SimpanData As String = "Insert into tblMasterPerkiraan values ('" & txtNoPerkiraan.Text & "','" & txtNamaPerkiraan.Text & "','" & cboGroup.Text & "','" & cboKeterangan.Text & "')"
                Command = New OleDbCommand(SimpanData, CONN)
                Command.ExecuteNonQuery()
                MsgBox("Data berhasil ditambahkan", MsgBoxStyle.Information, "")
                IsiList()
                KondisiAwal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Buka()
        Try
            If cmdEdit.Text = "Edit" Then

                cmdEdit.Text = "Update"
                cmdTambah.Enabled = False
                cmdSimpan.Enabled = False
                cmdHapus.Enabled = False
                cmdKeluar.Text = "Batal"
                cboGroup.Enabled = False
                txtNoPerkiraan.Enabled = False
                txtNamaPerkiraan.Focus()
            ElseIf txtNoPerkiraan.Text = "" Or txtNamaPerkiraan.Text = "" Or cboGroup.Text = "" Or cboKeterangan.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "")
            Else
                KoneksiKeAccess()
                Dim EditData As String = "Update tblMasterPerkiraan set NamaPerkiraan='" & txtNamaPerkiraan.Text & "',Kelompok='" & cboGroup.Text & "',Keterangan='" & cboKeterangan.Text & "' where NoPerkiraan= '" & txtNoPerkiraan.Text & "'"
                Command = New OleDbCommand(EditData, CONN)
                Command.ExecuteNonQuery()
                MsgBox("Data berhasil di Update", MsgBoxStyle.Information, "")
                IsiList()
                KondisiAwal()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdHapus_Click(sender As Object, e As EventArgs) Handles cmdHapus.Click
        Try
            Dim ans As String
            ans = MsgBox("Anda Yakin Ingin Menghapus?", MsgBoxStyle.YesNo, "Konfirmasi")
            If ans = vbYes Then
                KoneksiKeAccess()
                Dim HapusData As String = "Delete from tblMasterPerkiraan where NoPerkiraan= '" & txtNoPerkiraan.Text & "'"
                Command = New OleDbCommand(HapusData, CONN)
                Command.ExecuteNonQuery()
                MsgBox("Data berhasil di Hapus", MsgBoxStyle.Information, "")
                IsiList()
                KondisiAwal()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdTambah_Click(sender As Object, e As EventArgs) Handles cmdTambah.Click
        Buka()
        Try
            BersihkanIsian()
            IsiList()
            cmdEdit.Enabled = False
            cmdHapus.Enabled = False
            cmdSimpan.Enabled = True
            cmdKeluar.Text = "Batal"
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cboGroup_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboGroup.KeyPress
        If e.KeyChar = Chr(13) Then
            If cboGroup.Text = "" Then
                MsgBox("Kelompok / Group perkiraan tidak boleh kosong", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Pesan")
                cboGroup.Focus()
            Else
                cboKeterangan.Focus()
            End If
        End If
    End Sub

    Private Sub cboKeterangan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboKeterangan.KeyPress
        If (e.KeyChar = Chr(13)) Then
            If (cboKeterangan.Text = "") Then
                MsgBox("Keterangan perkiraan tidak boleh kosong", MsgBoxStyle.Critical, "Error")
            Else
                If (cmdSimpan.Enabled = True) Then
                    cmdSimpan.Focus()
                Else
                    cmdEdit.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub ListView_Click(sender As Object, e As EventArgs) Handles ListView.Click
        Tutup()
        cmdTambah.Enabled = False
        cmdSimpan.Enabled = False
        cmdHapus.Enabled = True
        cmdEdit.Enabled = True
        cmdKeluar.Text = "Batal"
    End Sub

    Private Sub ListView_DoubleClick(sender As Object, e As EventArgs) Handles ListView.DoubleClick
        Try
            frmAktivaTetap.lblNama.Text = ListView.SelectedItems.Item(0).SubItems(0).Text & " : " & ListView.SelectedItems.Item(0).SubItems(1).Text
            frmAktivaTetap.CariData()
            frmAktivaTetap.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ListView_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ListView.KeyPress
        If e.KeyChar = Chr(13) Then
            txtNamaPerkiraan.Focus()
        End If
    End Sub

    Private Sub ListView_KeyUp(sender As Object, e As KeyEventArgs) Handles ListView.KeyUp
        If e.KeyCode = Keys.Escape Then
            BersihkanIsian()
        End If
    End Sub

    Private Sub ListView_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView.SelectedIndexChanged
        txtNoPerkiraan.Enabled = False
        txtNamaPerkiraan.Enabled = True
        AmbilDataDariList()
    End Sub

    Private Sub cboGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboGroup.SelectedIndexChanged
        If Me.cboGroup.Text = "1-AKTIVA" Then
            'txtNoPerkiraan.Text = 1
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        ElseIf Me.cboGroup.Text = "2-HUTANG" Then
            'txtNoPerkiraan.Text = 2
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        ElseIf Me.cboGroup.Text = "3-MODAL" Then
            'txtNoPerkiraan.Text = 3
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        ElseIf Me.cboGroup.Text = "4-PENDAPATAN" Then
            'txtNoPerkiraan.Text = 4
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        ElseIf Me.cboGroup.Text = "5-PEMBELIAN" Then
            'txtNoPerkiraan.Text = 5
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        ElseIf Me.cboGroup.Text = "6-BIAYA" Then
            'txtNoPerkiraan.Text = 6
            txtNoPerkiraan.Enabled = True
            txtNoPerkiraan.Focus()
        End If
    End Sub

    Private Sub cmdPreview_Click(sender As Object, e As EventArgs) Handles cmdPreview.Click
        Try
            'frmRptCetakAkun.CrystalReportViewer1.SelectionFormula = "{tmpSaldoBlnLalu.Periode} = '" & cboPeriode.Text & "'"
            frmRptCetakAkun.CrystalReportViewer1.Dock = DockStyle.Fill
            frmRptCetakAkun.CrystalReportViewer1.RefreshReport()
            frmRptCetakAkun.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
End Class

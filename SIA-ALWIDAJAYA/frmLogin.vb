﻿Imports System.Data.OleDb
Public Class frmLogin

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtkode.Text = "ADM001"
        txtSandi.Text = "admin"
        txtkode.Focus()
    End Sub

    Private Sub cmdMasuk_Click(sender As Object, e As EventArgs) Handles cmdMasuk.Click
        Try
            If txtkode.Text = "" Or txtSandi.Text = "" Then
                MsgBox("Silahkan Masukkan Nama Pengguna dan Kata Sandi!", MsgBoxStyle.Information, "")
            Else
                KoneksiKeAccess()
                Command = New OleDbCommand("select * from tblMasterUser where KodeUser = '" & txtkode.Text & "' and KataKunci ='" & txtSandi.Text & "'", CONN)
                Rd = Command.ExecuteReader
                Rd.Read()
                If Rd.HasRows Then
                    Me.Close()
                    FormMenuUtama.Show()
                    'FormMenuJualBeli.lblAdminJ.Text = Rd.Item("NamaPengguna")
                    'FormMenuJualBeli.lblAdminT.Text = Rd.Item("NamaPengguna")
                    FormMenuUtama.txtName.Text = Rd.Item("NamaUser")
                    'FormMenuUtama.status = Rd.Item("level")
                    'If FormMenuUtama.status = "ADMIN" Then
                    '    FormMenuUtama.btnPengguna.Enabled = True
                    '    FormMenuUtama.btnCoA.Enabled = True
                    '    FormMenuUtama.btnPeriode.Enabled = True
                    'FormMenuUtama.btnBarang.Enabled = True
                    'FormMenuUtama.btnSupplier.Enabled = True
                    'FormMenuUtama.btnJualBeli.Enabled = True
                    'FormMenuUtama.btnJurnalUmum.Enabled = True
                    'FormMenuUtama.btnAJP.Enabled = True
                    'FormMenuUtama.btnProses.Enabled = True
                    'FormMenuUtama.btnLapAkun.Enabled = True
                    'FormMenuUtama.btnLapJU.Enabled = True
                    'FormMenuUtama.btnLapAJP.Enabled = True
                    'FormMenuUtama.btnLapNL.Enabled = True
                    'FormMenuUtama.btnLapPM.Enabled = True
                    'FormMenuUtama.btnLapRL.Enabled = True
                    'FormMenuUtama.btnLapNeraca.Enabled = True
                Else
                    'FormMenuUtama.btnPengguna.Enabled = False
                    'FormMenuUtama.btnCoA.Enabled = False
                    'FormMenuUtama.btnPeriode.Enabled = False
                    'FormMenuUtama.btnJurnalUmum.Enabled = False
                    'FormMenuUtama.btnAJP.Enabled = False
                    'FormMenuUtama.btnProses.Enabled = False
                    'FormMenuUtama.btnLapAkun.Enabled = False
                    'FormMenuUtama.btnLapJU.Enabled = False
                    'FormMenuUtama.btnLapAJP.Enabled = False
                    'FormMenuUtama.btnLapNL.Enabled = False
                    'FormMenuUtama.btnLapPM.Enabled = False
                    'FormMenuUtama.btnLapRL.Enabled = False
                    'FormMenuUtama.btnLapNeraca.Enabled = False
                    'FormMenuUtama.btnJualBeli.Enabled = True
                    'FormMenuUtama.btnBarang.Enabled = True
                    'FormMenuUtama.btnSupplier.Enabled = True
                End If
                'Else
                MsgBox("Kode Admin atau Password Salah!", MsgBoxStyle.Information, "")
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtkode_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtkode.KeyPress
        If e.KeyChar = Chr(13) Then
            txtSandi.Focus()
        End If
    End Sub
End Class